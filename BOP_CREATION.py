import xml.etree.ElementTree as gfg
from xml.dom import minidom
import pandas as pd


def generateXML(filename):
    # Read in input data from excel file
    file = pd.read_excel(filename, sheet_name=None, dtype=str, keep_default_na = False)

    # Iterate through all sheets of the excel file and create BOP for sheets with name "..._BOP"

    for sheetname, data in file.items():
        if sheetname[-3:] == "BOP":
            # Initial setup for BOP file
            global groupitem
            root = gfg.Element("BOP", attrib={'Name': sheetname})

            ps = gfg.Element("ProcessSegment", attrib={'ShortDesc': '', 'Name': ''})
            root.append(ps)
            pp = gfg.Element("ProcessParameters")
            ps.append(pp)
            rp = gfg.Element("RecipeParameters")
            ps.append(rp)

            # Iterate through every row of data
            for index, row in data.iterrows():

                # index = 0: create the first MCF
                if index == 0:
                    groupitem = createMCF(id=data.iloc[index, 4], parent=rp, index=data.iloc[index, 5],
                                          name=data.iloc[index, 6], description=data.iloc[index, 7],
                                          wftype=data.iloc[index, 8],
                                          source=data.iloc[index, 9], equipment=data.iloc[index, 10],
                                          sequence=data.iloc[index, 11], msg_group=data.iloc[index, 12], step_number=data.iloc[index, 13])
                # for subsequent item, create new MCF if current row is different from previous row
                elif data.iloc[index, 4] != data.iloc[index - 1, 4]:
                    groupitem = createMCF(data.iloc[index, 4], rp, data.iloc[index, 5], data.iloc[index, 6],
                                          data.iloc[index, 7], data.iloc[index, 8], data.iloc[index, 9],
                                          data.iloc[index, 10],
                                          data.iloc[index, 11], data.iloc[index, 12], step_number=data.iloc[index, 13])

                createOperation(id=data.iloc[index, 14], parent=groupitem, index=data.iloc[index, 15],
                                name=data.iloc[index, 16], description=data.iloc[index, 17],
                                iteration=data.iloc[index, 18])

                # if index > 1 and data.iloc[index, 2] != data.iloc[index - 1, 2]:
                xmlstr = minidom.parseString(gfg.tostring(root)).toprettyxml(indent="\t")
                with open(sheetname + ".xml", "w") as f:
                    f.write(xmlstr)


# Function to create MCF
def createMCF(id, parent, index='XXX', name='XXX', description='XXX', wftype='XXX', source='XXX', equipment='XXX',
              sequence='XXX', msg_group='XXX', step_number='XXX'):
    return gfg.SubElement(parent, "GroupItems",
                          attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                  'WF_Type': wftype, 'Source': source, 'Equipment': equipment,
                                  'Sequence': sequence, 'End_Sequence': '', 'Msg_Group': msg_group,
                                  'Step_Number': step_number})


# Function to create Operation
def createOperation(id, parent, index='XXX', name='XXX', description='XXX', iteration='XXX'):
    indicator = id[:-2]
    if indicator == "OP_CHARGE_MATERIAL":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (SCAN_CHARGE / AUTO_CHARGE / SCAN_CHECK / WD_SCAN_CHARGE)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CHARGE_SEQUENCE',
                                                   'Description': 'Charge sequence list',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BYPASS_SUMMARY',
                                                   'Description': 'Bypass summary message (YES / NO)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_INVENTORY_REFERENCE',
                                                   'Description': 'Inventory container reference (Barcode)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_QUANTITY',
                                                   'Description': 'Quantity to be charged (AUTO_CHARGE)',
                                                   'Value': ''})
    elif indicator == "OP_CHARGE_MATERIAL_SCALE":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action - Weigh mode (POSITVE / NEGATIVE)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CHARGE_SEQUENCE',
                                                   'Description': 'Charge sequence list',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_SCALE_ID',
                                                   'Description': 'Scale ID, e.g. GBL()',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CONTAINER_TYPE',
                                                   'Description': 'Container type',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EQUIPMENT_ID',
                                                   'Description': 'Equipment ID, e.g. GBL()',
                                                   'Value': 'XXX'})
    elif indicator == "OP_CHARGE_MATERIAL_TOTE":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (TOTE_CHARGE / TOTE_CHECK)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CHARGE_SEQUENCE',
                                                   'Description': 'Charge sequence list',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EQUIPMENT_ID',
                                                   'Description': 'Equipment ID, e.g. GBL()',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TOTE_STATION_ID',
                                                   'Description': 'Tote Station ID, e.g. GBL()',
                                                   'Value': 'XXX'})
    elif indicator == "OP_CONSUMABLE_CHECKS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (SCAN_ALLOCATE / SCAN_CHARGE / ALLOCATED_CHARGE / NON_BOM)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CHARGE_SEQUENCE',
                                                   'Description': 'Charge sequence list',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILTER_NUMBER_REQ',
                                                   'Description': 'Record filter serial number (YES / NO)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_QUANTITY',
                                                   'Description': 'Quantity',
                                                   'Value': ''})
    elif indicator == "OP_DATA_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (CALC / ADD DATE / DELTA DATE)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EXPRESSION',
                                                   'Description': 'Expression to evaluate (Formula)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1, e.g. Start Date',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2, e.g. End Date/ Interval',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM3',
                                                   'Description': 'Parameter 3, e.g. Interval UOM(Days, Hours, Minutes, Seconds, Months, Years)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM4',
                                                   'Description': 'Parameter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM5',
                                                   'Description': 'Parameter 5',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ROUND_DIGITS',
                                                   'Description': 'Round digits',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LIMIT_CHECK_ALIAS',
                                                   'Description': 'Limit check name',
                                                   'Value': ''})
    elif indicator == "OP_DV_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (CREATE_BATCH_UD / CREATE_BATCH / WAIT_FOR_DCS /  ADVANCE_DCS)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_DCS_ALIAS',
                                                   'Description': 'DeltaV alias, e.g. GBL(PRODUCTCODE)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILTER1',
                                                   'Description': 'Source equipment / filter 1',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILTER2',
                                                   'Description': 'Destination equipment / filter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILTER3',
                                                   'Description': 'Source equipment / filter 3',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILTER4',
                                                   'Description': 'Destination equipment / filter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_STEP_NUMBER',
                                                   'Description': 'Step Number',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM_VALUES',
                                                   'Description': 'Parameter Values to write (to advance DCS)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'String Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'String Parameter 2',
                                                   'Value': ''})
    elif indicator == "OP_ET_CHECKS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (AUTO / SELECT / SCAN)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EQUIPMENT_ALIAS',
                                                   'Description': 'Equipment alias from Process Parameter',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ET_CP_NAMES',
                                                   'Description': 'Custom property names to verify',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ET_CP_VALUES',
                                                   'Description': 'Custom property values to verify',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENTS_TO_PERFORM',
                                                   'Description': 'Events to perform (Delimiter: ~)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENT_PARAM_NAMES',
                                                   'Description': 'Event parameters names (Event1Par1|Event1Par2~Event2Par1)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENT_PARAM_VALUES',
                                                   'Description': 'Event parameters values (Event1Value1|Event1Value2|Event1Value3~Event2Value1|Event2Value2)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PROPERTIES_TO_DISPLAY',
                                                   'Description': 'Property names to display (Delimiter: |)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BYPASS_SUMMARY',
                                                   'Description': 'Bypass summary message (YES / NO)',
                                                   'Value': ''})
    elif indicator == "OP_ET_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action selection (Custom Property ReadWrite / ET Perform Event)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EQUIPMENT_ID',
                                                   'Description': 'Equipment, e.g. GBL()',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ET_CP_NAMES',
                                                   'Description': 'ET custom property names',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ET_CP_VALUES',
                                                   'Description': 'ET custom property values',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENT_ACTION',
                                                   'Description': 'Event Action (Write)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENTS_TO_PERFORM',
                                                   'Description': 'Events list for current equipment (Delimiter: ~)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENT_PARAM_NAMES',
                                                   'Description': 'List of required event parameters names (Event1Par1|Event1Par2|Event1Par3~Event2Par1|Event2Par2)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EVENT_PARAM_VALUES',
                                                   'Description': 'List of required event parameters values (Event1Value1|Event1Value2|Event1Value3~Event2Value1|Event2Value2)',
                                                   'Value': ''})
    elif indicator == "OP_FILLING_1":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILLING_ALIAS',
                                                   'Description': 'Filling alias, e.g. BAG_FILLING_1',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BAG_NUMBER_START',
                                                   'Description': 'Bag number start',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_SCALE_ID',
                                                   'Description': 'Scale ID',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PRINTER_REFERENCE',
                                                   'Description': 'Printer reference',
                                                   'Value': 'SHF Area'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BYPASS_OUTPUT',
                                                   'Description': 'Bypass output creation (YES / NO)',
                                                   'Value': 'NO'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EXPIRY_ALIAS',
                                                   'Description': 'Expiry Alias',
                                                   'Value': 'OUTPUT_EXPIRY'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_OPC_ALIAS',
                                                   'Description': 'OPC tag alias',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_DCS_DVI_ALIAS',
                                                   'Description': 'DCS DVI alias',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILL_START_STEP_NO',
                                                   'Description': 'Filling start (Initial weight) step number',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILL_END_STEP_NO',
                                                   'Description': 'Filling end (Final weight) step number',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_NEXT_BAG_STEP_NO',
                                                   'Description': 'Next bag step number',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
    elif indicator == "OP_FILLING_2":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_OUTPUT_FUNCTION',
                                                   'Description': 'Output function (CREATE_OUTPUT / CREATE_OUTPUT_RTI)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FILLING_ALIAS',
                                                   'Description': 'Filling alias, e.g. BOTTLE_FILLING_1',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EXPIRY_ALIAS',
                                                   'Description': 'Expiry Alias',
                                                   'Value': 'OUTPUT_EXPIRY'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BOTTLE_NUMBER_START',
                                                   'Description': 'Bottle number start',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PRINTER_REFERENCE',
                                                   'Description': 'Printer reference',
                                                   'Value': 'SHF Area'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TOTAL_QTY_UOM',
                                                   'Description': 'Total filled quantity UOM',
                                                   'Value': 'XXX'})
    elif indicator == "OP_MISC_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (DATETIME_ENTRY / ATTACH_FILES_UD / ATTACH_FILES / PRINT_LABEL)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_NUMBER_OF_LABELS',
                                                   'Description': 'Number of labels',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LABEL_TYPE',
                                                   'Description': 'Label Type',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_DATETIME_FORMAT',
                                                   'Description': 'Date time format',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ATTACHMENT_TYPE',
                                                   'Description': 'Attachment Type, e.g. Sample Results',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PRINTER_REFERENCE',
                                                   'Description': 'Printer Reference',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM_NAME',
                                                   'Description': 'Input Parameter Name',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'String Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'String Parameter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM3',
                                                   'Description': 'String Parameter 3',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM4',
                                                   'Description': 'String Parameter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM5',
                                                   'Description': 'String Parameter 5',
                                                   'Value': ''})
    elif indicator == "OP_OM_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (WD / AD_INVENTORY / CHARGE_SCRAP)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CHARGE_SEQUENCE',
                                                   'Description': 'Charge sequence list',
                                                   'Value': 'XXX'})
    elif indicator == "OP_OPC_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_SOURCE',
                                                   'Description': 'Source',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Function (READ_AUTO)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FORMAT',
                                                   'Description': 'Format of output value',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_OPC_TAG_ALIAS',
                                                   'Description': 'OPC tag alias',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LIMIT_CHECK_ALIAS',
                                                   'Description': 'Limit check alias',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_BYPASS_SUMMARY',
                                                   'Description': 'Bypass summary message',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'String Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'String Parameter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM3',
                                                   'Description': 'String Parameter 3',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM4',
                                                   'Description': 'String Parameter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM5',
                                                   'Description': 'String Parameter 5',
                                                   'Value': ''})
    elif indicator == "OP_OUTPUT_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action to perform (Create Output RTI / Output RTI)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PRINTER_REFERENCE',
                                                   'Description': 'Printer',
                                                   'Value': 'SHF Area'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LOCATION',
                                                   'Description': 'Location barcode (-$DSP)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_CONTAINER_COUNT',
                                                   'Description': 'Number of output containers',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_OUTPUT_QTY',
                                                   'Description': 'Output quantity',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_OUTPUT_UOM',
                                                   'Description': 'BOM UOM',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EXPIRY_DATE',
                                                   'Description': 'Expiry date',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_EXPIRY_ALIAS',
                                                   'Description': 'Expiry Alias',
                                                   'Value': 'OUTPUT_EXPIRY'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2',
                                                   'Value': ''})
    elif indicator == "OP_SAMPLE_RESULT_1":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TEST_NAME',
                                                   'Description': 'Test name (pH / Conductivity)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TEMPERATURE_REQ',
                                                   'Description': 'Temperature required (YES / NO)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TEST_LIMIT_ALIAS',
                                                   'Description': 'Test limit check name',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TEMP_LIMIT_ALIAS',
                                                   'Description': 'Temperature limit check name',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2',
                                                   'Value': ''})
    elif indicator == "OP_SAMPLING":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action to be performed (SAP Sampling / PU Sampling / Sample Submission)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_SAMPLE_ALIAS',
                                                   'Description': 'Sample Alias / Name in BOP Process Parameters section',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TRANSFER_LOCATION',
                                                   'Description': 'Sample transfer location',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
    elif indicator == "OP_SCALE_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action - Scale function (Zero)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LIMIT_CHECK_ALIAS',
                                                   'Description': 'Limit check name',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_SCALE_ID',
                                                   'Description': 'Scale ID',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM3',
                                                   'Description': 'Parameter 3',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM4',
                                                   'Description': 'Parameter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM5',
                                                   'Description': 'Parameter 5',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_FORMAT',
                                                   'Description': 'Output value format',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_UOM',
                                                   'Description': 'Scale UOM',
                                                   'Value': ''})
    elif indicator == "OP_SCAN_PARTS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PART_ALIAS',
                                                   'Description': 'Part (Item) alias',
                                                   'Value': 'XXX'})
    elif indicator == "OP_UI_FUNCTIONS":
        group = gfg.SubElement(parent, "Group",
                               attrib={'ID': id, 'Index': index, 'Name': name, 'Description': description,
                                       'Iteration': iteration})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ACTION',
                                                   'Description': 'Action (PROMPT / SELECT / ENTRY / TIMER)',
                                                   'Value': 'XXX'})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_LIMIT_CHECK_ALIAS',
                                                   'Description': 'Limit check name',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_ENUM_LIST',
                                                   'Description': 'Enumeration list (CSV)',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM_NAME',
                                                   'Description': 'Input parameter name',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM1',
                                                   'Description': 'Parameter 1',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM2',
                                                   'Description': 'Parameter 2',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM3',
                                                   'Description': 'Parameter 3',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM4',
                                                   'Description': 'Parameter 4',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM5',
                                                   'Description': 'Parameter 5',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TIMER_INTERVAL',
                                                   'Description': 'Timer interval',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_TIMER_UOM',
                                                   'Description': 'Timer interval UOM',
                                                   'Value': ''})
        gfg.SubElement(group, "Parameter", attrib={'Name': 'OP_RI_PARAM_FORMAT',
                                                   'Description': 'Round digits',
                                                   'Value': ''})
    else:
        print("{} does not exist!".format(indicator))


# Driver Code
if __name__ == "__main__":
    generateXML('HKA_SHF_Toan.xlsm')
