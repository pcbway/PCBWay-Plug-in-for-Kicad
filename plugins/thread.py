#copyright  Aisler and licensed under the MIT license.
#https://opensource.org/licenses/MIT 

import os
import webbrowser
import shutil
import json
import requests
import datetime
import re
import wx
import tempfile
origcwd = os.getcwd()
os.chdir(os.path.split(__file__)[0])
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
os.chdir(origcwd)
from threading import Thread
from .result_event import *
from .config import *


def processFootprint(FPID):
  try:
    text = str(FPID.GetFootprintName())
  except AttributeError:
    text = str(FPID.GetLibItemName())
  m = re.fullmatch(r"[^:]*:(.*)", text)
  if m:
    text = m.group(1)
  m = re.fullmatch(r"[A-Za-z]+_([0-9]+)_[0-9]+Metric", text)
  if m:
    # we use the SMD case code based on inches, this is the generally used package size code
    text = "_" + m.group(1) # prefix case code with underscore because later imports with excel would eat the leading zero, which is a dangerous error!
  return text

def processPartSideName(name):
  if name == "F.Cu":
    return "top"
  elif name == "B.Cu":
    return "bottom"
  else:
    return name

def parseReference(text):
  m = re.fullmatch(r"([A-Za-z]+)([0-9]+)", text)
  if m:
    return (m.group(1), int(m.group(2)))
  else:
    return (text, 0)

def getPartValue(footprint):
  value = footprint.GetValueAsString()
  value = value.replace(" ", "_") # space cannot be represented in the plain text format!
  value = value.replace("\t", "_")
  return value

class PCBWayThread(Thread):
    def __init__(self, wxObject):
        Thread.__init__(self)
        self.wxObject = wxObject
        self.start()

    def generatePositionfile(self, board):
      origin = board.GetDesignSettings().GetAuxOrigin()
      if origin.x == 0 and origin.y == 0:
        wx.MessageBox('Missing Aux origin point from PCB', 'Error', wx.OK | wx.ICON_INFORMATION)
        raise RuntimeError("Missing origin")
      fileName = os.path.splitext(board.GetFileName())[0] + ".pos"
      columnNames = ["Ref", "Val", "Package", "PosX", "PosY", "Rot", "Side"]
      lines = []
      for footprint in board.Footprints():
        ref = footprint.GetReferenceAsString()
        refp = parseReference(ref)
        value = getPartValue(footprint)
        package = processFootprint(footprint.GetFPID())
        pos = footprint.GetPosition()
        pos -= origin
        posX = float(pos.x) / pcbnew.PCB_IU_PER_MM
        posY = -float(pos.y) / pcbnew.PCB_IU_PER_MM
        deg = footprint.GetOrientationDegrees()
        side = processPartSideName(footprint.GetLayerName())
        if footprint.GetAttributes() & pcbnew.FP_SMD != pcbnew.FP_SMD: # parts without any SMD pads are excluded from the position file (purely throughhole part)
          continue
        if value == "DNP": #old way of marking it DNP
          continue
        if footprint.HasProperty("dnp"): # new flag in KiCAD 7.0.0
          continue
        if (footprint.GetAttributes() & pcbnew.FP_EXCLUDE_FROM_POS_FILES) == pcbnew.FP_EXCLUDE_FROM_POS_FILES: # KiCAD 6.0.0 preferred way of marking DNP
          continue
        lines.append((ref, value, package, posX, posY, deg, side, refp))
      
      lines = sorted(lines, key=lambda x: (x[6], x[7]))
      columnSizes = [4, 6, 6, 10, 10, 10, 10]
      
      for row in lines:
        for i,data in enumerate(row[:len(columnSizes)]):
          if type(data) is str:
            columnSizes[i] = max(columnSizes[i], len(data)+2)
          elif type(data) is float:
            columnSizes[i] = max(columnSizes[i], len("%.4f"%(data))+2)
      
      columnSizes[0] *= -1
      columnSizes[1] *= -1
      columnSizes[2] *= -1
      columnSizes[6] *= -1
      
      with open(fileName, "wt") as f:
        now = datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
        f.write("### Footprint positions - created on %s ### \n"%(now))
        f.write("### Printed by PCBWay plugin\n")
        f.write("## Unit = mm, Angle = deg.\n")
        f.write("## Side : All\n")
        f.write("# ")
        for i,name in enumerate(columnNames):
          if i == 0:
            offset = 2
          else:
            offset = 0
          f.write(("%%%is"%(columnSizes[i]+offset))%name + " ")
        f.write("\n")
        
        for row in lines:
          assert row[0].find(" ") == -1
          assert row[1].find(" ") == -1
          assert row[2].find(" ") == -1
          assert row[6].find(" ") == -1
          f.write(("%%%is %%%is %%%is %%%i.4f %%%i.4f %%%i.4f %%%is\n"%tuple(columnSizes))%row[:7])
        f.write("## End\n")

    def generateBOM(self, board):
      fileName = os.path.splitext(board.GetFileName())[0] + "-bom.xlsx"
      columnNames = ["Item #", "Ref Des", "Qty", "Manufacturer", "Mfg part #", "Description / Value", "Package", "Type"]
      footprints = []
      for footprint in board.Footprints():
        ref = footprint.GetReferenceAsString()
        refp = parseReference(ref)
        value = getPartValue(footprint)
        package = processFootprint(footprint.GetFPID())
        props = footprint.GetProperties()
        manufacturer = props.get("Manufacturer_Name", props.get("Manufacturer Name", props.get("Manufacturer", "")))
        manufacturerPN = props.get("Manufacturer_Part_Number", props.get("Manufacturer Part Number", props.get("Part Number", "")))
        if footprint.GetAttributes() & (pcbnew.FP_SMD | pcbnew.FP_THROUGH_HOLE) == pcbnew.FP_SMD:
          groupType = "SMD"
        if footprint.GetAttributes() & (pcbnew.FP_SMD | pcbnew.FP_THROUGH_HOLE) == pcbnew.FP_THROUGH_HOLE:
          groupType = "THT"
        if footprint.GetAttributes() & (pcbnew.FP_SMD | pcbnew.FP_THROUGH_HOLE) == (pcbnew.FP_THROUGH_HOLE | pcbnew.FP_SMD):
          groupType = "SMD&THT"
        if package.startswith("TestPoint_Pad"): # we do not need parts for a bare test pad on the PCB
          continue
        if value == "DNP": #old way of marking it DNP
          continue
        if footprint.HasProperty("dnp"): # new flag in KiCAD 7.0.0
          groupType = "DNP"
        if (footprint.GetAttributes() & pcbnew.FP_EXCLUDE_FROM_BOM) == pcbnew.FP_EXCLUDE_FROM_BOM: # KiCAD 6.0.0 preferred way of marking DNP
          continue
        bomkey = (value, package, groupType, manufacturer, manufacturerPN)
        footprints.append({"ref" : ref, "refp" : refp, "value" : value, "package" : package, "footprint" : footprint, "bomkey" : bomkey, "manufacturer" : manufacturer, "manufacturerPN" : manufacturerPN, "type" : groupType})
      footprints = sorted(footprints, key=lambda x: x["refp"])
      groups = {}
      for data in footprints:
        groups[data["bomkey"]] = groups.get(data["bomkey"], []) + [data]
      groups = sorted([groups[key] for key in groups], key=lambda x:x[0]["ref"])
      
      workbook = xlsxwriter.Workbook(fileName)
      worksheet = workbook.add_worksheet()

      wrap_format = workbook.add_format()
      wrap_format.set_text_wrap()
      wrap_format.set_border(1)

      wrap_format_warning = workbook.add_format()
      wrap_format_warning.set_text_wrap()
      wrap_format_warning.set_border(1)
      wrap_format_warning.set_bg_color("yellow")

      wrap_format_warning2 = workbook.add_format()
      wrap_format_warning2.set_text_wrap()
      wrap_format_warning2.set_border(1)
      wrap_format_warning2.set_bg_color("#FFAF08")
      
      headerFormat = workbook.add_format()
      headerFormat.set_bold()
      headerFormat.set_bg_color("gray")
      headerFormat.set_border(1)
      
      columnSizes = [8, 40, 8, 20, 20, 30, 30, 10]
      for i,value in enumerate(columnSizes):
        worksheet.set_column(i, i, value)
      
      worksheet.write(0, 1, "Board filename:")
      worksheet.write(0, 2, board.GetFileName())
      
      row_longest = {}
      
      headerRow = 7
      
      worksheet.set_row(headerRow, 30)
      
      for colindex, data in enumerate(columnNames):
        worksheet.write(headerRow, colindex, data, headerFormat)
      
      #for i,rowindex in enumerate(range(headerRow+1, headerRow+1+len(groups))):
        #if i == 0:
          #worksheet.write(rowindex, 0, 1, wrap_format)
        #else:
          #worksheet.write_formula(rowindex, 0, "=" + xl_rowcol_to_cell(rowindex-1, 0) + "+1", value=i+1, cell_format=wrap_format)
      
      row_longest = {}
      
      tht_num = 0
      hybrid_num = 0
      smd_num = 0
      
      for i,group in enumerate(groups):
        rowindex = i + headerRow + 1
        refs = [data["ref"] for data in group]
        refs = ", ".join(refs)
        row = [refs, len(group), group[0]["value"], group[0]["manufacturer"], group[0]["manufacturerPN"], group[0]["package"], group[0]["type"]]
        if group[0]["type"] == "SMD":
          smd_num += len(group)
        if group[0]["type"] == "SMD&THT":
          hybrid_num += len(group)
        if group[0]["type"] == "THT":
          tht_num += len(group)
        if group[0]["type"] == "DNP":
          cell_format = wrap_format_warning
        elif group[0]["type"] == "THT":
          cell_format = wrap_format_warning2
        else:
          cell_format = wrap_format
        if i == 0:
          worksheet.write(rowindex, 0, 1, cell_format)
        else:
          worksheet.write_formula(rowindex, 0, "=OFFSET(" + xl_rowcol_to_cell(rowindex, 0) + ",-1,0) + 1", value=i+1, cell_format=cell_format)
        for j,value in enumerate(row):
          row_longest[rowindex] = max(row_longest.get(rowindex, 0), (len(str(value))+columnSizes[j+1]-1) // columnSizes[j+1])
          worksheet.write(rowindex, j+1, value, cell_format)
      
      for rowindex in row_longest:
        worksheet.set_row(rowindex, max(max(row_longest[rowindex], 1)*12.0, 13.0) + 3)
      
      worksheet.write(3, 1, "SMD part count:")
      worksheet.write_formula(3, 2, '=SUMIFS('+xl_rowcol_to_cell(headerRow+1, 2)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 2)+', '+xl_rowcol_to_cell(headerRow+1, 7)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 7)+',"SMD")', value=smd_num)
      
      worksheet.write(4, 1, "Through Hole part count:")
      worksheet.write_formula(4, 2, '=SUMIFS('+xl_rowcol_to_cell(headerRow+1, 2)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 2)+', '+xl_rowcol_to_cell(headerRow+1, 7)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 7)+',"THT")', value=tht_num)
      
      if hybrid_num:
        worksheet.write(5, 1, "Hybrid part count:")
        worksheet.write_formula(5, 2, '=SUMIFS('+xl_rowcol_to_cell(headerRow+1, 2)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 2)+', '+xl_rowcol_to_cell(headerRow+1, 7)+':'+xl_rowcol_to_cell(headerRow+1+len(groups)*2, 7)+',"SMD&THT")', value=hybrid_num)
      
      workbook.close()
        
    
    def run(self):
        
        temp_dir = tempfile.mkdtemp()
        _, temp_file = tempfile.mkstemp()
        board = pcbnew.GetBoard()
        title_block = board.GetTitleBlock()
        p_name = board.GetFileName()

        self.report(5)

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
        if hasattr(popt, "SetExcludeEdgeLayer"):
            popt.SetExcludeEdgeLayer(True)
        popt.SetUseGerberProtelExtensions(False)
        popt.SetUseAuxOrigin(True)
        popt.SetSubtractMaskFromSilk(False)
        popt.SetDrillMarksType(0)  # NO_DRILL_SHAPE

        self.report(20)
        for layer_info in plotPlan:
            if board.IsLayerEnabled(layer_info[1]):
                pctl.SetLayer(layer_info[1])
                pctl.OpenPlotfile(
                    layer_info[0],
                    pcbnew.PLOT_FORMAT_GERBER,
                    layer_info[2])
                pctl.PlotLayer()

        pctl.ClosePlot()

        self.report(25)
        drlwriter = pcbnew.EXCELLON_WRITER(board)

        drlwriter.SetOptions(
            False,
            True,
            board.GetDesignSettings().GetAuxOrigin(),
            False)
        drlwriter.SetFormat(False)
        drlwriter.CreateDrillandMapFilesSet(pctl.GetPlotDirName(), True, False)
        
        self.report(30)
        netlist_writer = pcbnew.IPC356D_WRITER(board)
        netlist_writer.Write(os.path.join(temp_dir, netlistFilename))
        
        self.report(35)
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

            attrs = f.GetAttributes()
            parsed_attrs = self.parse_attrs(attrs)

            if f.HasProperty("dnp"): # new attribute in KiCad 7.0, the value of the attribute does not matter in the 7.0.0 version
                parsed_attrs["not_in_pos"] = True
                parsed_attrs["not_in_bom"] = True

            mount_type = 'smt' if parsed_attrs['smd'] else 'tht'
            placed = not parsed_attrs['not_in_bom']

            rotation = f.GetOrientation().AsDegrees() if hasattr(f.GetOrientation(), 'AsDegrees') else f.GetOrientation() / 10.0
            designator = f.GetReference()

            components.append({
                'pos_x': (f.GetPosition()[0] - board.GetDesignSettings().GetAuxOrigin()[0]) / 1000000.0,
                'pos_y': (f.GetPosition()[1] - board.GetDesignSettings().GetAuxOrigin()[1]) * -1.0 / 1000000.0,
                'rotation': rotation,
                'side': layer,
                'designator': designator,
                'mpn': self.getMpnFromFootprint(f),
                'pack': footprint_name,
                'value': f.GetValue(),
                'mount_type': mount_type,
                'place': placed
            })
        
        boardWidth = pcbnew.ToMM(board.GetBoardEdgesBoundingBox().GetWidth())
        boardHeight = pcbnew.ToMM(board.GetBoardEdgesBoundingBox().GetHeight())
        if hasattr(board, 'GetCopperLayerCount'):
            boardLayer = board.GetCopperLayerCount()
        with open((os.path.join(temp_dir, componentsFilename)), 'w') as outfile:
            json.dump(components, outfile)

        self.generatePositionfile(board)
        try:
          self.generateBOM(board)
        except Exception as e:
          wx.MessageBox('Error:' + str(e), 'Error', wx.OK | wx.ICON_INFORMATION)
        #
        
        return
      
        temp_file = shutil.make_archive(p_name, 'zip', temp_dir)
        files = {'upload[file]': open(temp_file, 'rb')}

        upload_url = baseUrl + '/Common/KiCadUpFile/'
        
        self.report(45)
        
        rsp = requests.post(
            upload_url, files=files, data={'boardWidth':boardWidth,'boardHeight':boardHeight,'boardLayer':boardLayer})
        
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
                self.report(45 + percent / 1.8)

        webbrowser.open(urls['redirect'])
        self.report(-1)

    def report(self, status):
        wx.PostEvent(self.wxObject, ResultEvent(status))
        
    def getMpnFromFootprint(self, f):
        keys = ['mpn', 'MPN', 'Mpn', 'PCBWay_MPN']
        for key in keys:
            if f.HasProperty(key):
                return f.GetProperty(key)

    def parse_attr_flag(self, attr, mask):
        return mask == (attr & mask)

    def parse_attrs(self, attrs):
        return {} if not isinstance(attrs, int) else {
            'tht': self.parse_attr_flag(attrs, pcbnew.FP_THROUGH_HOLE),
            'smd': self.parse_attr_flag(attrs, pcbnew.FP_SMD),
            'not_in_pos': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_POS_FILES),
            'not_in_bom': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_BOM),
            'not_in_plan': self.parse_attr_flag(attrs, pcbnew.FP_BOARD_ONLY)
        }
    
