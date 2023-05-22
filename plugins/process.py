#https://opensource.org/licenses/MIT 

import os
import json
import csv
import re
from .config import *


class PCBWayProcess:
    def __init__(self):
        self.board = pcbnew.GetBoard()
        self.pctl = pcbnew.PLOT_CONTROLLER(self.board)
        self.bom = []
        self.components = []

    def get_gerber_file(self, temp_dir):
        settings = self.board.GetDesignSettings()
        settings.m_SolderMaskMargin = 0
        settings.m_SolderMaskMinWidth = 0

        #pctl = pcbnew.PLOT_CONTROLLER(self.board)

        popt = self.pctl.GetPlotOptions()
        popt.SetOutputDirectory(temp_dir)
        popt.SetPlotFrameRef(False)
        popt.SetSketchPadLineWidth(pcbnew.FromMM(0.1))
        popt.SetAutoScale(False)
        popt.SetScale(1)
        popt.SetMirror(False)
        popt.SetUseGerberAttributes(True)
        if hasattr(popt, "SetExcludeEdgeLayer"):
            popt.SetExcludeEdgeLayer(True)
        popt.SetUseGerberProtelExtensions(False)
        popt.SetUseAuxOrigin(True)
        popt.SetSubtractMaskFromSilk(True)
        popt.SetDrillMarksType(0)  # NO_DRILL_SHAPE

        for layer_info in plotPlan:
            if self.board.IsLayerEnabled(layer_info[1]):
                self.pctl.SetLayer(layer_info[1])
                self.pctl.OpenPlotfile(
                    layer_info[0],
                    pcbnew.PLOT_FORMAT_GERBER,
                    layer_info[2])
                self.pctl.PlotLayer()

        self.pctl.ClosePlot()

    def get_netlist_file(self, temp_dir):
        drlwriter = pcbnew.EXCELLON_WRITER(self.board)

        drlwriter.SetOptions(
            False,
            True,
            self.board.GetDesignSettings().GetAuxOrigin(),
            False)
        drlwriter.SetFormat(False)
        drlwriter.CreateDrillandMapFilesSet(self.pctl.GetPlotDirName(), True, False)
        
        netlist_writer = pcbnew.IPC356D_WRITER(self.board)
        netlist_writer.Write(os.path.join(temp_dir, netlistFilename))

    def get_components_file(self, temp_dir):
        if hasattr(self.board, 'GetModules'):
            footprints = list(self.board.GetModules())
        else:
            footprints = list(self.board.GetFootprints())

        footprints.sort(key=lambda x: x.GetReference())
        
        for i, f in enumerate(footprints):
            try:
                footprint_name = str(f.GetFPID().GetFootprintName())
            except AttributeError:
                footprint_name = str(f.GetFPID().GetLibItemName())

            layer = {
                pcbnew.F_Cu: 'top',
                pcbnew.B_Cu: 'bottom',
            }.get(f.GetLayer())

            f_attrs = f.GetAttributes()
            parsed_attrs = self.parse_attrs(f_attrs)

            mount_type = 'smt' if parsed_attrs['smd'] else 'tht'
            placed = not parsed_attrs['not_in_bom']

            rotation = f.GetOrientation().AsDegrees() if hasattr(f.GetOrientation(), 'AsDegrees') else f.GetOrientation() / 10.0
            designator = f.GetReference()

            pos_x = (f.GetPosition()[0] - self.board.GetDesignSettings().GetAuxOrigin()[0]) / 1000000.0
            pos_y = (f.GetPosition()[1] - self.board.GetDesignSettings().GetAuxOrigin()[1]) * -1.0 / 1000000.0

            mpn = self.get_mpn_from_footprint(f)

            value = f.GetValue()
            self.components.append({
                'pos_x': pos_x,
                'pos_y': pos_y,
                'rotation': rotation,
                'side': layer,
                'designator': designator,
                'mpn': mpn,
                'pack': footprint_name,
                'value': value,
                'mount_type': mount_type,
                'place': placed
            })

            is_exist_bom = False

            for exist_bom in self.bom:
                if exist_bom['mpn'] == mpn and exist_bom['pack'] == footprint_name and exist_bom['value'] == value:
                    exist_bom['designator'] += ', ' + designator
                    exist_bom['quantity'] += 1
                    is_exist_bom = True

            if is_exist_bom == False:
                self.bom.append({
                    'designator': designator,
                    'quantity': 1,
                    'value': value,
                    'pack': footprint_name,
                    'mpn': mpn,
                    'mount_type': mount_type
                })
        
       
        if len(self.components) > 0:
            with open((os.path.join(temp_dir, positionsFilename)), 'w', newline='', encoding='utf-8-sig') as outfile:
                csvobj = csv.writer(outfile)
                csvobj.writerow(self.components[0].keys())

                for component in self.components:
                    if ('**' not in component['designator']):
                        csvobj.writerow(component.values())

        if len(self.bom) > 0:
            with open((os.path.join(temp_dir, bomFilename)), 'w', newline='', encoding='utf-8-sig') as outfile:
                csvobj = csv.writer(outfile)
                csvobj.writerow(self.bom[0].keys())

                for component in self.bom:
                    if ('**' not in component['designator']):
                        csvobj.writerow(component.values())


    def get_gerber_parameter(self):
        
        boardWidth = pcbnew.ToMM(self.board.GetBoardEdgesBoundingBox().GetWidth())
        boardHeight = pcbnew.ToMM(self.board.GetBoardEdgesBoundingBox().GetHeight())

        if hasattr(self.board, 'GetCopperLayerCount'):
            boardLayer = self.board.GetCopperLayerCount()
        gerberData = {'boardWidth':boardWidth,'boardHeight':boardHeight,'boardLayer':boardLayer}
        return gerberData
    
    def get_name(self):

        p_name = self.board.GetFileName()

        return p_name

    def parse_attrs(self, attrs):
        return {} if not isinstance(attrs, int) else {
            'tht': self.parse_attr_flag(attrs, pcbnew.FP_THROUGH_HOLE),
            'smd': self.parse_attr_flag(attrs, pcbnew.FP_SMD),
            'not_in_pos': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_POS_FILES),
            'not_in_bom': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_BOM),
            'not_in_plan': self.parse_attr_flag(attrs, pcbnew.FP_BOARD_ONLY)
        }
        
    def get_mpn_from_footprint(self, f):
        keys = ['mpn', 'MPN', 'Mpn', 'PCBWay_MPN' ,'part number', 'Part Number', 'Part No.', 'Mfr. Part No.', 'Mfg Part']
        for key in keys:
            if f.HasProperty(key):
                return f.GetProperty(key)

    def parse_attr_flag(self, attr, mask):
        return mask == (attr & mask)


        
