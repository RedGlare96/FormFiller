import logging
import os
from os import path
from sys import stdout
from datetime import datetime
from configparser import ConfigParser
import pdfrw
import openpyxl


def check_create_dir(dirname):
    '''
    Checks if directory exists and if it doesn't creates a new directory
    :param dirname: Path to directory
    '''
    if not path.exists(dirname):
        if '/' in dirname:
            os.makedirs(dirname)
        else:
            os.mkdir(dirname)


def fill_form(targetfile, outputfile, row_data, mapping):
    '''
    Fill up form based on map data
    :param targetfile: Path to pdf file that needs to be filled
    :param outputfile: Path to output
    :param row_data: Row data from excel file
    :param mapping: Dictionary that maps excel row with form field in pdf
    '''
    logger = logging.getLogger(__name__ + '.fill_form')
    template = pdfrw.PdfFileReader(targetfile)
    for page in range(len(template.pages)):
        annotations = template.pages[page]['/Annots']
        for item in annotations:
            if item['/Subtype'] == '/Widget':
                if item['/T']:
                    key = item['/T'][1: -1]
                    if key in mapping.keys():
                        if 'Check Box' in key:
                            if row_data[mapping[key]] == 'yes':
                                item.update(pdfrw.PdfDict(AS=pdfrw.PdfName('Yes')))
                            elif row_data[mapping[key]] == 'no':
                                item.update(pdfrw.PdfDict(AS=pdfrw.PdfName('No')))
                        else:
                            logger.debug('Textbox= {0}  Value= {1}'.format(str(key), str(row_data[mapping[key]])))
                            item.update(pdfrw.PdfDict(V=str(row_data[mapping[key]])))
    logger.info('Writing to {}'.format(path.join(OUTPUT_FOLDER, outputfile)))
    pdfrw.PdfWriter().write(path.join(OUTPUT_FOLDER, outputfile), template)


def update_form(row_data, mapping, targetfile=None, template=None):
    '''
    Fills up form and returns a pdfrw PdfileWriter object. Works with an existing object if specified. If not, creates new
    :param row_data: Row data from excel file
    :param mapping: Dictionary that maps excel row with form field in pdf
    :param targetfile: Path to pdf file that needs to be filled
    :param template:
    :return:
    '''
    logger = logging.getLogger(__name__ + '.update_form')
    if template is None:
        if targetfile is not None:
            template = pdfrw.PdfReader(targetfile)
        else:
            raise Exception('Target file not included. Cannot create template')
    for page in range(len(template.pages)):
        annotations = template.pages[page]['/Annots']
        for item in annotations:
            if item['/Subtype'] == '/Widget':
                if item['/T']:
                    key = item['/T'][1: -1]
                    if key in mapping.keys():
                        if 'Check Box' in key:
                            if row_data[mapping[key]] == 'yes':
                                item.update(pdfrw.PdfDict(AS=pdfrw.PdfName('Yes')))
                            elif row_data[mapping[key]] == 'no':
                                item.update(pdfrw.PdfDict(AS=pdfrw.PdfName('No')))
                        else:
                            logger.debug('Textbox= {0}  Value= {1}'.format(str(key), str(row_data[mapping[key]])))
                            item.update(pdfrw.PdfDict(V=str(row_data[mapping[key]])))
    return template


def get_rows(input_file, worksheet):
    '''
    Extracts excel data as a dictionary with column name as key
    :param input_file: Path to excel file
    :param worksheet: Worksheet name to get data from
    :return: Dictionary with column name as key and row data as value
    '''
    wb = openpyxl.load_workbook(input_file)
    ws = wb[worksheet]
    cols = []
    ret = []
    for i, row in enumerate(ws.iter_rows()):
        new_row = dict()
        for j, ele in enumerate(row):
            if i == 0:
                try:
                    cols.append(ele.value.lower())
                except AttributeError:
                    cols.append('extra data')
            else:
                new_row[cols[j]] = ele.value
        if i > 0:
            ret.append(new_row)
    return ret


def invert_dict(d):
    '''
    Inverts key and value in dictionaries
    :param d: Input dictionaries
    :return: Inverted dictionary
    '''
    return dict([(v, k) for k, v in d.items()])


if __name__ == '__main__':
    print('Formfiller')
    print('Reading config...')
    # init config
    config = ConfigParser()
    config.read('masterconfig.ini')
    OUTPUT_FOLDER = config['paths']['output_folder']
    map_dir = config['paths']['map_dir']
    defaults_dir = config['paths']['defaults_dir']
    check_create_dir(OUTPUT_FOLDER)

    # Init logging
    rootLogger = logging.getLogger()
    rootLogger.setLevel(logging.DEBUG)
    consoleHandler = logging.StreamHandler(stdout)
    consoleHandler.setFormatter(logging.Formatter('[%(name)s] - %(levelname)s - %(message)s'))
    check_create_dir('logs')
    fileHandler = logging.FileHandler(
        path.join('logs', 'Formfiller{0}.log'.format(datetime.now().strftime('%d-%m-%y-%H-%M-%S'))))
    fileHandler.setFormatter(logging.Formatter('%(asctime)s:-[%(name)s] - %(levelname)s - %(message)s'))
    consoleHandler.setLevel(logging.INFO)
    rootLogger.addHandler(consoleHandler)
    fileHandler.setLevel(logging.DEBUG)
    rootLogger.addHandler(fileHandler)

    for mapfile in os.listdir(map_dir):
        datamap = path.join(map_dir, mapfile)
        mapconfig = ConfigParser()
        mapconfig.read(datamap)
        # Get settings
        source_file = mapconfig['settings']['source_file']
        form = mapconfig['settings']['pdf_form']
        identifier = mapconfig['settings']['identifier']
        base_map = mapconfig['settings']['base_map']
        # Init values
        econ_units_list = []
        # Iterate through worksheet sections
        for ele in mapconfig._sections.keys():
            if ele != 'settings' and ele != 'econ units':
                # Discards settings and special case
                defaultmap = None
                defaultvalues = None
                worksheetname = ele
                rootLogger.info('Processing {}:'.format(worksheetname))
                map = mapconfig._sections[worksheetname]

                # Look for defaults
                if path.exists(path.join(defaults_dir, '{}-defaults.ini'.format(worksheetname))):
                    rootLogger.debug('Defaults file found')
                    defaultconfig = ConfigParser()
                    defaultconfig.read(path.join(defaults_dir, '{}-defaults.ini'.format(worksheetname)))
                    defaultmap = defaultconfig._sections['mapping']
                    defaultvalues = defaultconfig._sections['values']

                # Iterate through rows and write into seperate pdf files
                for row in get_rows(source_file, worksheetname):
                    # Add default values if the exist
                    if defaultvalues is not None:
                        final_values = {**row, **defaultvalues}
                    else:
                        final_values = row
                    if defaultmap is not None:
                        final_map = invert_dict({**map, **defaultmap})
                    else:
                        final_map = invert_dict(map)
                    outputfile = '{}.pdf'.format(final_values[identifier])
                    fill_form(form, outputfile, final_values, final_map)

            # Handle special case
            elif ele == 'econ units':
                # Init variables
                instance_cnt = 0
                attachment_cnt = 0
                prev_identifier = 'initvalue'
                base_template = None
                outputname = 'If-you-see-this-its-error'
                unit_templates = []
                final_map = None
                final_values = None
                new_slot = None
                defaultmap = None
                defaultvalues = None
                worksheetname = ele
                rootLogger.info('Processing {}:'.format(worksheetname))

                # Look for defaults
                if path.exists(path.join(defaults_dir, '{}-defaults.ini'.format(worksheetname))):
                    print('Defaults file found')
                    defaultconfig = ConfigParser()
                    defaultconfig.read(path.join(defaults_dir, '{}-defaults.ini'.format(worksheetname)))
                    defaultmap = defaultconfig._sections['mapping']
                    defaultvalues = defaultconfig._sections['values']

                # Iterate through row data
                for row in get_rows(source_file, worksheetname):
                    # Get identifier
                    new_identifier = row[identifier]
                    # Check with previous identifier
                    if new_identifier == prev_identifier:
                        # If the raw has the same identifier, it is another instance
                        instance_cnt += 1
                        rootLogger.debug('Update instance {}'.format(instance_cnt))
                    else:
                        # New property found. Pool PREVIOUS collected values and save all of them
                        rootLogger.debug('New property')
                        # Begin saving PREVIOUS base template
                        if base_template is not None:
                            outputfile = '{}.pdf'.format(outputname)
                            rootLogger.info('Writing {}'.format(outputfile))
                            pdfrw.PdfWriter().write(path.join(OUTPUT_FOLDER, outputfile), base_template)
                            if new_slot is not None:
                                # Append remaining slots to unit_templates
                                unit_templates.append(new_slot)
                            if len(unit_templates) > 0:
                                # Iterate through collected attachments and save all of them
                                for i, item in enumerate(unit_templates):
                                    rootLogger.info('Saving attachment {}'.format(i))
                                    rootLogger.info('Writing: {}'.format('{0}-A{1}.pdf'.format(outputname, i)))
                                    pdfrw.PdfWriter().write(path.join(OUTPUT_FOLDER, '{0}-A{1}.pdf'.format(outputname, i)),
                                                            item)
                        # Re-initialize values by de-referencing them
                        instance_cnt = 0
                        attachment_cnt = 0
                        base_template = None
                        unit_templates = list()
                        new_slot = None
                    prev_identifier = new_identifier
                    if instance_cnt > 4:
                        # All 4 slots in the econ unit page is over. Another page required
                        # Collect current attachment into list
                        unit_templates.append(new_slot)
                        # Re-initialize slot reference
                        new_slot = None
                        attachment_cnt += 1
                        instance_cnt = 1
                        rootLogger.debug(
                            'Update attachment - attachment: {0} instance: {1}'.format(attachment_cnt, instance_cnt))

                    # Begin instance insertions
                    if instance_cnt == 0:
                        # 1st instance of property is added to base template
                        map = mapconfig._sections[base_map]
                        # Add default values if they exist
                        if defaultvalues is not None:
                            final_values = {**row, **defaultvalues}
                        else:
                            final_values = row
                        if defaultmap is not None:
                            final_map = invert_dict({**map, **defaultmap})
                        else:
                            final_map = invert_dict(map)
                        new_targetfile = form
                        base_template = update_form(final_values, final_map, targetfile=new_targetfile)
                        outputname = final_values[identifier]
                    else:
                        rootLogger.debug('Instance: {}'.format(instance_cnt))
                        final_map = {value: key.split(';')[0]
                                     for key, value in mapconfig._sections[worksheetname].items() if
                                     ';{}'.format(instance_cnt) in key}
                        final_values = row
                        new_targetfile = mapconfig['settings']['pdf_form2']
                        if instance_cnt == 1:
                            # If slot is initialized, create new template
                            rootLogger.debug('Creating new slot')
                            new_slot = update_form(final_values, final_map, targetfile=new_targetfile)
                        else:
                            # If slot already exists, update existing
                            rootLogger.debug('Updating slot')
                            new_slot = update_form(final_values, final_map, template=new_slot)

                # The last set of instances are saved outside
                # the iterations as the identifier change triggers saving routine of the previous set of instances
                rootLogger.debug('Last property')
                if base_template is not None:
                    outputfile = '{}.pdf'.format(outputname)
                    rootLogger.info('Writing {}'.format(outputfile))
                    pdfrw.PdfWriter().write(path.join(OUTPUT_FOLDER, outputfile), base_template)
                    if new_slot is not None:
                        # Append remaining slots to unit_templates
                        unit_templates.append(new_slot)
                    if len(unit_templates) > 0:
                        for i, item in enumerate(unit_templates):
                            rootLogger.info('Saving attachment {}'.format(i))
                            rootLogger.info('Writing: {}'.format('{0}-A{1}.pdf'.format(outputname, i)))
                            pdfrw.PdfWriter().write(path.join(OUTPUT_FOLDER, '{0}-A{1}.pdf'.format(outputname, i)), item)


