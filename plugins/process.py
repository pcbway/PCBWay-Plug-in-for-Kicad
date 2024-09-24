#https://opensource.org/licenses/MIT 

import pcbnew # type: ignore
import os
import csv

from itertools import groupby

from .config import *
from .utils import *

class PCBWayProcess:
    def __init__(self):
        self.board = pcbnew.GetBoard()
        self.pctl = pcbnew.PLOT_CONTROLLER(self.board)
        self.bom = []
        self.components = []
    
    def get_name(self):
        return self.board.GetFileName()
        
    def get_basedir(self):
        return os.path.dirname(self.board.GetFileName())

    def get_basename(self):
        return os.path.basename(self.board.GetFileName())
    
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
        
        mpn_keys = get_mpn_keys()
        pack_keys = get_pack_keys()
        no_show_keys = [
            'ki_fp_filters',
            'DNP',
            'Reference',
            'Value',
            'Datasheet',
            'Footprint',
        ]
        ignore_ext_keys = mpn_keys + pack_keys + no_show_keys

        greater_v8 = is_greater_v8()
        fp_datas = []
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
            not_in_bom = parsed_attrs['not_in_bom']
            not_in_pos = parsed_attrs['not_in_pos']

            if not_in_bom and not_in_pos:
                continue

            rotation = f.GetOrientation().AsDegrees() if hasattr(f.GetOrientation(), 'AsDegrees') else f.GetOrientation() / 10.0

            pos_x = (f.GetPosition()[0] - self.board.GetDesignSettings().GetAuxOrigin()[0]) / 1000000.0
            pos_y = (f.GetPosition()[1] - self.board.GetDesignSettings().GetAuxOrigin()[1]) * -1.0 / 1000000.0

            designator = f.GetReference()
            value = f.GetValue()
            mpn = get_mpn_from_footprint(f)
            pack = get_pack_from_footprint(f)
            is_dnp = get_is_dnp_from_footprint(f) if greater_v8 else False

            if not footprint_name:
                footprint_name = ''

            if not pack:
                pack = ''

            if not mpn:
                mpn = ''

            if not_in_pos == False:
                self.components.append({
                    'pos_x': pos_x,
                    'pos_y': pos_y,
                    'rotation': rotation,
                    'side': layer,
                    'designator': designator,
                    'mpn': mpn,
                    'pack': pack,
                    'footprint': footprint_name,
                    'value': value,
                    'mount_type': mount_type,
                })

            if not_in_bom:
                continue
            
            fp_item_fields = {
                'designator': designator,
                'value': value,
                'footprint': footprint_name,
                'pack': pack,
                'mpn': mpn,
                'DNP': 'Yes' if is_dnp else '',
                'Mount_Type': mount_type,
            }

            if greater_v8:
                footprint_fields = f.GetFieldsText()
                if footprint_fields:
                    for k, v in footprint_fields.items():
                        if k.upper() == 'DNP' or k.upper() == 'MOUNT_TYPE':
                            k = 'Custom_' + k
                        if not v or k in ignore_ext_keys:
                            continue
                        fp_item_fields[k] = v

            fp_datas.append(fp_item_fields)

        fp_data_group = {}
        for item in fp_datas:
            designator = item['designator']
            value = item['value']
            footprint = item['footprint']
            pack = item['pack']
            mpn = item['mpn']
            is_dnp = item['DNP']

            index = value + '_' + footprint + '_' + pack + '_' + mpn
            if is_dnp:
                index = designator + '_' + index

            if index in fp_data_group:
                fp_data_group[index].append(item)
            else:
                fp_data_group[index] = [ item ]

        fixed_columns = [
            'designator',
            'quantity',
            'value',
            'footprint',
            'pack',
            'mpn',
        ]

        all_columns = [
            'Designator',
            'Quantity',
            'Value',
            'Footprint',
            'Package',
            'MPN',
        ]
        rows = []
        for _key, items in fp_data_group.items():
            first_item = items[0]

            row_datas = {}
            row_columns = []
            designators = []
            for item in items:
                designator = item['designator']
                designators.append(designator)

                for item_key, item_value in item.items():
                    if item_key in fixed_columns:
                        continue
                    if item_key not in all_columns:
                        all_columns.append(item_key)
                    if item_key not in row_columns:
                        row_columns.append(item_key)
                    if item_key not in row_datas:
                        row_datas[item_key] = []
                    row_datas[item_key].append({
                        'key': designator,
                        'value': item_value,
                    })

            for k in row_columns:
                row_data = row_datas[k]
                row_data_groupby = {val: list(group) for val, group in groupby(row_data, key=lambda x: x['value'])}
                item_text = ''
                if len(row_data_groupby) > 1:
                    item_text = '; '.join(['[' + ','.join([g['key'] for g in group]) + ']'+ group_key for group_key, group in row_data_groupby.items()])
                else:
                    item_text = row_data[0]['value']
                first_item[k] = item_text

            designator_count = len(designators)
            designator = ', '.join(designators)
            

            row = {
                'Designator': designator,
                'Quantity': designator_count,
                'Value': first_item['value'],
                'Footprint': first_item['footprint'],
                'Package': first_item['pack'],
                'MPN': first_item['mpn'],
            }

            for k in row_columns:
                if k in first_item:
                    row[k] = first_item[k]

            rows.append(row)
        
        for row in rows:
            newRow = {}
            for k in all_columns:
                if k in row:
                    newRow[k] = row[k]
                else:
                    newRow[k] = ''
            if not greater_v8:
                del newRow['DNP']
            self.bom.append(newRow)
       
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
                    if ('**' not in component['Designator']):
                        csvobj.writerow(component.values())


    def get_gerber_parameter(self):
        boardWidth = pcbnew.ToMM(self.board.GetBoardEdgesBoundingBox().GetWidth())
        boardHeight = pcbnew.ToMM(self.board.GetBoardEdgesBoundingBox().GetHeight())

        if hasattr(self.board, 'GetCopperLayerCount'):
            boardLayer = self.board.GetCopperLayerCount()

        return {
            'boardWidth':boardWidth,
            'boardHeight':boardHeight,
            'boardLayer':boardLayer,
        }

    def parse_attrs(self, attrs):
        return {} if not isinstance(attrs, int) else {
            'tht': self.parse_attr_flag(attrs, pcbnew.FP_THROUGH_HOLE),
            'smd': self.parse_attr_flag(attrs, pcbnew.FP_SMD),
            'not_in_pos': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_POS_FILES),
            'not_in_bom': self.parse_attr_flag(attrs, pcbnew.FP_EXCLUDE_FROM_BOM),
            'not_in_plan': self.parse_attr_flag(attrs, pcbnew.FP_BOARD_ONLY)
        }

    def parse_attr_flag(self, attr, mask):
        return mask == (attr & mask)
