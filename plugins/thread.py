import os
import webbrowser
import shutil
import json
import requests
import re
import wx
import time
import tempfile
from threading import Thread
from .result_event import *
from .config import *

#from requests_toolbelt.multipart import encoder


class PCBWayThread(Thread):
    def __init__(self, wxObject):
        Thread.__init__(self)
        self.wxObject = wxObject
        self.start()

    def run(self):
        temp_dir = tempfile.mkdtemp()
        _, temp_file = tempfile.mkstemp()
        board = pcbnew.GetBoard()
        title_block = board.GetTitleBlock()
        self.report(10)
        match = re.match(
            '^PCBWay Project ID: ([A-Z]{8})$',
            title_block.GetComment(commentLineIdx))
        if match:
            project_id = match.group(1)
        else:
            project_id = False

        # Override a few design parameters as our CAM takes care of this
        settings = board.GetDesignSettings()
        settings.m_SolderMaskMargin = 0
        settings.m_SolderMaskMinWidth = 0

        pctl = pcbnew.PLOT_CONTROLLER(board)

        popt = pctl.GetPlotOptions()
        popt.SetOutputDirectory(temp_dir)
        popt.SetPlotFrameRef(False)
        popt.SetSketchPadLineWidth(pcbnew.FromMM(0.1))
        popt.SetAutoScale(False)
        popt.SetScale(1)
        popt.SetMirror(False)
        popt.SetUseGerberAttributes(True)
        popt.SetExcludeEdgeLayer(True)
        popt.SetUseGerberProtelExtensions(False)
        popt.SetUseAuxOrigin(True)
        popt.SetSubtractMaskFromSilk(False)
        popt.SetDrillMarksType(0)  # NO_DRILL_SHAPE

        self.report(15)
        for layer_info in plotPlan:
            if board.IsLayerEnabled(layer_info[1]):
                pctl.SetLayer(layer_info[1])
                pctl.OpenPlotfile(
                    layer_info[0],
                    pcbnew.PLOT_FORMAT_GERBER,
                    layer_info[2])
                pctl.PlotLayer()

        pctl.ClosePlot()

        # Write excellon drill files
        self.report(20)
        drlwriter = pcbnew.EXCELLON_WRITER(board)

        # mirrot, header, offset, mergeNPTH
        drlwriter.SetOptions(
            False,
            True,
            board.GetDesignSettings().GetAuxOrigin(),
            False)
        drlwriter.SetFormat(False)
        drlwriter.CreateDrillandMapFilesSet(pctl.GetPlotDirName(), True, False)

        # # Write netlist to enable Smart Tests
        self.report(25)
        netlist_writer = pcbnew.IPC356D_WRITER(board)
        netlist_writer.Write(os.path.join(temp_dir, netlistFilename))

        # # Export component list
        self.report(30)
        components = []
        if hasattr(board, 'GetModules'):
            footprints = list(board.GetModules())
        else:
            footprints = list(board.GetFootprints())

        for i, f in enumerate(footprints):
            try:
                footprint_name = str(f.GetFPID().GetFootprintName())
            except AttributeError:
                footprint_name = str(f.GetFPID().GetLibItemName())

            layer = {
                pcbnew.F_Cu: 'top',
                pcbnew.B_Cu: 'bottom',
            }.get(f.GetLayer())

            mount_type = {
                0: 'smt',
                1: 'tht',
                2: 'smt'
            }.get(f.GetAttributes())

            components.append({
                'pos_x': (f.GetPosition()[0] - board.GetDesignSettings().GetAuxOrigin()[0]) / 1000000.0,
                'pos_y': (f.GetPosition()[1] - board.GetDesignSettings().GetAuxOrigin()[1]) * -1.0 / 1000000.0,
                'rotation': f.GetOrientation() / 10.0,
                'side': layer,
                'designator': f.GetReference(),
                'mpn': self.getMpnFromFootprint(f),
                'pack': footprint_name,
                'value': f.GetValue(),
                'mount_type': mount_type
            })

        boardWidth = pcbnew.Iu2Millimeter(board.GetBoardEdgesBoundingBox().GetWidth())
        boardHeight = pcbnew.Iu2Millimeter(board.GetBoardEdgesBoundingBox().GetHeight())
        boardLayer = board.GetCopperLayerCount()
        
        with open((os.path.join(temp_dir, componentsFilename)), 'w') as outfile:
            json.dump(components, outfile)

        # # Create ZIP file
        temp_file = shutil.make_archive(temp_file, 'zip', temp_dir)
        files = {'upload[file]': open(temp_file, 'rb')}

        upload_url = baseUrl + '/Common/KiCadUpFile/'
        try:
            self.report(40)
            if project_id:
                data = {}
                data['upload_url'] = baseUrl + '/Common/KiCadUpFile/'
                data['project_id'] = project_id
            else:
                rsp = requests.get(baseUrl + '/Common/NewKiCad/')
                data = json.loads(rsp.content)
                if not title_block.GetComment(commentLineIdx):
                    title_block.SetComment(
                        commentLineIdx,
                        'PCBWay Project ID: ' +
                        data['project_id'])

            rsp = requests.post(
                upload_url, files=files, data={'project_id':data['project_id'],'boardWidth':boardWidth,'boardHeight':boardHeight,'boardLayer':boardLayer})
            
            urls = json.loads(rsp.content)

            readsofar = 0
            totalsize = os.path.getsize(temp_file)
            with open(temp_file, 'rb') as file:
                while True:
                    data = file.read(10)
                    if not data:
                        break
                    readsofar += len(data)
                    percent = readsofar * 1e2 / totalsize
                    self.report(40 + percent / 1.7)

            # progress = 0
            # while progress < 100:
            #     time.sleep(pollingInterval)
            #     progress = json.loads(
            #         requests.get(
            #             urls['callback']).content)['progress']
            #     self.report(40 + progress / 1.7)

            webbrowser.open(urls['redirect'])
            self.report(-1)
        except Exception as e:
            webbrowser.open(repr(e))

    def report(self, status):
        wx.PostEvent(self.wxObject, ResultEvent(status))
        
    def getMpnFromFootprint(self, f):
        keys = ['mpn', 'MPN', 'Mpn', 'AISLER_MPN']
        for key in keys:
            if f.HasProperty(key):
                return f.GetProperty(key)
    