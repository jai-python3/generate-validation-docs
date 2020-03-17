import csv
from docx import Document
from docx.shared import Inches
import os
import sys
import click
import pathlib
import json
import logging
import calendar
import time
import pathlib
from colorama import Fore, Style
from datetime import datetime
from mailmerge import MailMerge
from datetime import date

DEFAULT_TEMPLATE_FILES_DIR = os.path.dirname(os.path.abspath(__file__)) + '/validation_documents/'

IQ_TEMPLATE_FILE_BASENAME = 'IQ_Checklist_template.docx'
OQ_TEMPLATE_FILE_BASENAME = 'OQ_Validation_Testing_Worksheet_template.docx'
PQ_TEMPLATE_FILE_BASENAME = 'PQ_Validation_Testing_Worksheet_template.docx'
SYSTEM_SPECIFICATION_TEMPLATE_FILE_BASENAME = 'System_Specification_template.docx'
TEST_PLAN_TEMPLATE_FILE_BASENAME = 'Test_Plan_template.docx'
USER_REQUIREMENTS_TEMPLATE_FILE_BASENAME = 'User_Requirements_template.docx'
VALIDATION_REPORT_TEMPLATE_FILE_BASENAME = 'Validation_Report_template.docx'

DEFAULT_document_prepared_by = 'Jaideep Sundaram'

DEFAULT_DOCUMENT_PREPARED_DATE = str(datetime.today().strftime('%d-%b-%Y'))

DEFAULT_OUTDIR = "/tmp/" + os.path.basename(__file__) + '/' + str(datetime.today().strftime('%Y-%m-%d-%H%M%S'))

DEFAULT_CONFIG_FILE = os.path.dirname(os.path.abspath(__file__)) + '/conf/config.json'

g_config = None
g_software_name = None
g_software_version = None
g_document_prepared_by = None
g_document_prepared_date = None
g_template_files_dir = None
g_outdir = None
g_server = None

g_iq_checklist_ctr = 0

g_iq_software_checklist_table_records = None
g_iq_hardware_checklist_table_records = None

g_iq_yes_no = None
g_iq_date = None

g_oq_yes_no = None
g_oq_date = None

g_pq_yes_no = None
g_pq_date = None

LOGGING_FORMAT = "%(levelname)s : %(asctime)s : %(pathname)s : %(lineno)d : %(message)s"

LOG_LEVEL = logging.INFO


def instantiate_mailmerge(template_file):
    """Instantiate the MailMerge class

    :param template_file: {str} the template file
    :return document: {Object} the MailMerge instance 
    """
    document = MailMerge(template_file)

    logging.info(document.get_merge_fields())

    document.merge(
        document_prepared_by=g_document_prepared_by,
        document_prepared_date=g_document_prepared_date,
        software_name=g_software_name,
        software_version=g_software_version,
        server=g_server)

    return document


def get_template_file(doc_type, default_template_file_basename=None):
    """Derive the template file from the config file
    :param doc_type: {str} document type
    :param default_template_file_basename: {str}
    :return template_file: {str}
    """
    template_file_basename = None

    if doc_type not in g_config or 'template file basename' not in g_config[doc_type]:
        template_file_basename = default_template_file_basename
        logging.info("Could not retrieve the '{} 'template file basename' from the config file so set default '{}'".format(doc_type, template_file_basename))
    else:
        template_file_basename = g_config[doc_type]['template file basename']

    template_file = os.path.join(g_template_files_dir, template_file_basename)

    if not os.path.exists(template_file):
        logging.error("template file '{}' does not exist".format(template_file))

    return template_file


def prepare_validation_document(template_file, outfile):
    """Prepare the specific validation document
    :param template_file: {str} the MS Word template file
    :param outfile: {str} the output file
    :return:
    """
    if not os.path.exists(template_file):
        logging.error("template file '{}' does not exist".format(template_file))

    document = instantiate_mailmerge(template_file)
    document.write(outfile)
    print("Wrote output file '{}'".format(outfile))


def get_iq_hardware_checklist_file(doc_type):
    """Derive the IQ hardware checklist file
    :param doc_type: {str} the document type, default 'IQ'
    :return infile: {str} the IQ hardware checklist file
    """
    if doc_type not in g_config or 'hardware checklist file basename' not in g_config[doc_type]:
        error_msg = "Could not retrieve the '{}' 'hardware checklist file basename' from the config file".format(doc_type)
        logging.error(error_msg)
        raise Exception(error_msg)
    else:
        basename = g_config[doc_type]['hardware checklist file basename']
        infile = os.path.dirname(os.path.abspath(__file__)) + '/conf/' + basename

    if not os.path.exists(infile):
        raise Exception("file '{}' does not exist".format(infile))

    return infile


def get_iq_hardware_table_records(doc_type='IQ'):
    """Parse the IQ hardware checklist tab-delimited file and build list of dictionaries
    :param doc_type: {str} document type, default 'IQ'
    :return hardware_table_records: {list} array of dictionaries
    """

    global g_iq_hardware_checklist_table_records
    global g_iq_checklist_ctr

    if g_iq_hardware_checklist_table_records is None:

        infile = get_iq_hardware_checklist_file(doc_type)

        header_to_position_lookup = {}
        record_ctr = 0

        hardware_table_records = []

        with open(infile) as f:
            reader = csv.reader(f, delimiter='\t')
            row_ctr = 0
            for row in reader:
                row_ctr += 1
                if row_ctr == 1:
                    field_ctr = 0
                    for field in row:
                        header_to_position_lookup[field] = field_ctr
                        field_ctr += 1
                    logging.info("Processed the header of csv file '{}'".format(infile))
                else:
                    g_iq_checklist_ctr += 1

                    description = row[header_to_position_lookup['Description']]
                    requirement = row[header_to_position_lookup['Requirement']]
                    record_lookup = {
                        'h_id': str(g_iq_checklist_ctr),
                        'h_desc': description,
                        'h_req': requirement,
                        'h_yes_no': g_iq_yes_no,
                        'h_date': g_iq_date
                    }
                    hardware_table_records.append(record_lookup)
                    record_ctr += 1
            logging.info("Processed '{}' records in tab-delimited file '{}'".format(record_ctr, infile))

        logging.info(hardware_table_records)
        g_iq_hardware_checklist_table_records = hardware_table_records

    return g_iq_hardware_checklist_table_records


def get_iq_software_checklist_file(doc_type):
    """Derive the IQ software checklist file
    :param doc_type: {str} the document type, default 'IQ'
    :return infile: {str} the IQ software checklist file
    """
    if doc_type not in g_config or 'software checklist file basename' not in g_config[doc_type]:
        error_msg = "Could not retrieve the '{}' 'software checklist file basename' from the config file".format(doc_type)
        logging.error(error_msg)
        raise Exception(error_msg)
    else:
        basename = g_config[doc_type]['software checklist file basename']
        infile = os.path.dirname(os.path.abspath(__file__)) + '/conf/' + basename

    if not os.path.exists(infile):
        raise Exception("file '{}' does not exist".format(infile))

    return infile


def get_iq_software_table_records(doc_type):
    """Parse the IQ software checklist tab-delimited file and build list of dictionaries
    :param doc_type: {str}
    :return software_table_records: {list} array of dictionaries
    """
    global g_iq_software_checklist_table_records
    global g_iq_checklist_ctr

    if g_iq_software_checklist_table_records is None:

        infile = get_iq_software_checklist_file(doc_type)

        header_to_position_lookup = {}
        record_ctr = 0

        software_table_records = []

        with open(infile) as f:
            reader = csv.reader(f, delimiter='\t')
            row_ctr = 0
            for row in reader:
                row_ctr += 1
                if row_ctr == 1:
                    field_ctr = 0
                    for field in row:
                        header_to_position_lookup[field] = field_ctr
                        field_ctr += 1
                    logging.info("Processed the header of csv file '{}'".format(infile))
                else:
                    g_iq_checklist_ctr += 1
                    description = row[header_to_position_lookup['Description']]
                    requirement = row[header_to_position_lookup['Requirement']]
                    software_table_records.append({
                        's_id': str(g_iq_checklist_ctr),
                        's_desc': description,
                        's_req': requirement,
                        's_yes_no': g_iq_yes_no,
                        's_date': g_iq_date
                    })
                    record_ctr += 1
            logging.info("Processed '{}' records in tab-delimited file '{}'".format(record_ctr, infile))

        logging.info(software_table_records)

        g_iq_software_checklist_table_records = software_table_records
    return g_iq_software_checklist_table_records


def prepare_iq():
    """Prepare the IQ Checklist validation document
    :return None:
    """
    doc_type = 'IQ'

    global g_iq_yes_no
    global g_iq_date

    yes_no = input("Prepare executed IQ? [Y/n] ")
    yes_no = yes_no.strip()
    if yes_no is None or yes_no == '' or yes_no == 'Y' or yes_no == 'y':
        g_iq_yes_no = 'Yes'
        g_iq_date = DEFAULT_DOCUMENT_PREPARED_DATE
        logging.info("Will prepare partially executed IQ validation document")
    elif yes_no == 'N' or yes_no == 'n':
        g_iq_yes_no = ''
        g_iq_date = ''
        logging.info("Will not prepare a partially executed IQ validation document")

    template_file = get_template_file(doc_type, IQ_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)

    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - IQ Checklist - ' + g_document_prepared_date + '.docx'

    hardware_table_records = get_iq_hardware_table_records(doc_type)
    software_table_records = get_iq_software_table_records(doc_type)

    document.merge_rows('h_id', hardware_table_records)
    document.merge_rows('s_id', software_table_records)

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def get_oq_checklist_file(doc_type='OQ'):
    """Derive the OQ checklist tab-delimited file
    :param doc_type: {str} the document type, default 'OQ'
    :return infile: {str} the absolute path to the OQ checklist tab-delimited file
    """
    if doc_type not in g_config or 'checklist file basename' not in g_config[doc_type]:
        error_msg = "Could not retrieve the '{}' 'checklist file basename' from the config file".format(doc_type)
        logging.error(error_msg)
        raise Exception(error_msg)
    else:
        basename = g_config[doc_type]['checklist file basename']
        infile = os.path.dirname(os.path.abspath(__file__)) + '/conf/' + basename

    if not os.path.exists(infile):
        raise Exception("'{}' checklist file '{}' does not exist".format(doc_type, infile))

    return infile


def get_oq_checklist_tables(doc_type='OQ'):
    """Retrieve the OQ checklist data for replicate 1 and replicate 2 from the tab-delimited file
    :param doc_type: {str} the document type default 'OQ'
    :return checklist_tables: {list} containing two arrays for each checklist replicate which in turn are arrays of dictionaries
    """

    infile = get_oq_checklist_file(doc_type)

    header_to_position_lookup = {}
    record_ctr = 0

    checklist_replicate1_table_records = []
    checklist_replicate2_table_records = []
    id_ctr = 0

    test_numbers_included = False

    with open(infile) as f:
        reader = csv.reader(f, delimiter='\t')
        row_ctr = 0
        for row in reader:
            row_ctr += 1
            if row_ctr == 1:
                field_ctr = 0
                for field in row:
                    if field == 'Test Number':
                        test_numbers_included = True
                    header_to_position_lookup[field] = field_ctr
                    field_ctr += 1
                logging.info("Processed the header of csv file '{}'".format(infile))
            else:
                id_ctr += 1
                if test_numbers_included:
                    test_id = row[header_to_position_lookup['Test Number']]
                else:
                    test_id = 'T' + str(id_ctr)

                checklist_replicate1_table_records.append({
                    'id_rep1': test_id,
                    'test_procedure_rep1': row[header_to_position_lookup['Test Procedure']],
                    'expected_finding_rep1': row[header_to_position_lookup['Expected Finding']],
                    'yes_no': g_oq_yes_no,
                    'date_initialed': g_oq_date
                })

                checklist_replicate2_table_records.append({
                    'id_rep2': test_id,
                    'test_procedure_rep2': row[header_to_position_lookup['Test Procedure']],
                    'expected_finding_rep2': row[header_to_position_lookup['Expected Finding']],
                    'yes_no': g_oq_yes_no,
                    'date_initialed': g_oq_date
                })

                record_ctr += 1

        logging.info("Processed '{}' records in tab-delimited file '{}'".format(record_ctr, infile))

    return [checklist_replicate1_table_records, checklist_replicate2_table_records]


def get_oq_test_data_file(doc_type='OQ'):
    """Derive the OQ test data tab-delimited file
    :param doc_type: {str} the document type, default 'OQ'
    :return infile: {str} the absolute path for the OQ test data tab-delimited file
    """
    infile = None

    if doc_type not in g_config or 'test data file basename' not in g_config[doc_type]:
        error_msg = "Could not retrieve the '{}' 'test data file basename' from the config file".format(doc_type)
        logging.warning(error_msg)
    else:
        basename = g_config[doc_type]['test data file basename']
        infile = os.path.dirname(os.path.abspath(__file__)) + '/conf/' + basename
        if not os.path.exists(infile):
            raise Exception("'{}' test data file '{}' does not exist".format(doc_type, infile))

    return infile


def get_oq_test_data_records(doc_type='OQ'):
    """Retrieve the OQ checklist data for replicate 1 and replicate 2 from the tab-delimited file
    :param doc_type: {str} the document type default 'OQ'
    :return checklist_tables: {list} containing two arrays for each checklist replicate which in turn are arrays of dictionaries
    """
    infile = get_oq_test_data_file(doc_type)

    test_data_records = []

    if infile is None:
        test_data_records = [{
            'test_data_name': 'TBD',
            'test_data_desc': 'TBD'
        }]
    else:
        header_to_position_lookup = {}
        record_ctr = 0

        with open(infile) as f:
            reader = csv.reader(f, delimiter='\t')
            row_ctr = 0
            for row in reader:
                row_ctr += 1
                if row_ctr == 1:
                    field_ctr = 0
                    for field in row:
                        header_to_position_lookup[field] = field_ctr
                        field_ctr += 1
                    logging.info("Processed the header of csv file '{}'".format(infile))
                else:

                    record_lookup = {
                        'test_data_name': row[header_to_position_lookup['Name']],
                        'test_data_desc': row[header_to_position_lookup['Description']]
                    }

                    test_data_records.append(record_lookup)

                    record_ctr += 1
            logging.info("Processed '{}' records in tab-delimited file '{}'".format(record_ctr, infile))

    return test_data_records


def prepare_oq():
    """Prepare the OQ Validation Testing Worksheet validation document
    :return None:
    """
    doc_type = 'OQ'

    global g_oq_yes_no
    global g_oq_date

    yes_no = input("Prepare executed OQ? [Y/n] ")
    yes_no = yes_no.strip()
    if yes_no is None or yes_no == '' or yes_no == 'Y' or yes_no == 'y':
        g_oq_yes_no = 'Yes'
        g_oq_date = DEFAULT_DOCUMENT_PREPARED_DATE
        logging.info("Will prepare partially executed OQ validation document")
    elif yes_no == 'N' or yes_no == 'n':
        g_oq_yes_no = ''
        g_oq_date = ''
        logging.info("Will not prepare a partially executed OQ validation document")

    template_file = get_template_file(doc_type, OQ_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)

    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - OQ Validation Testing Worksheet - ' + g_document_prepared_date + '.docx'

    checklist_tables = get_oq_checklist_tables()

    test_data_records = get_oq_test_data_records()

    document.merge_rows('test_data_name', test_data_records)
    document.merge_rows('id_rep1', checklist_tables[0])
    document.merge_rows('id_rep2', checklist_tables[1])

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def prepare_pq():
    """Prepare the PQ Validation Testing Worksheet validation document
    :return None:
    """
    doc_type = 'PQ'
    
    global g_pq_yes_no
    global g_pq_date

    yes_no = input("Prepare executed PQ? [Y/n] ")
    yes_no = yes_no.strip()
    if yes_no is None or yes_no == '' or yes_no == 'Y' or yes_no == 'y':
        g_pq_yes_no = 'Yes'
        g_pq_date = DEFAULT_DOCUMENT_PREPARED_DATE
        logging.info("Will prepare partially executed PQ validation document")
    elif yes_no == 'N' or yes_no == 'n':
        g_pq_yes_no = ''
        g_pq_date = ''
        logging.info("Will not prepare a partially executed PQ validation document")

    template_file = get_template_file(doc_type, PQ_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)

    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - PQ Validation Testing Worksheet - ' + g_document_prepared_date + '.docx'

    checklist_tables = get_oq_checklist_tables()

    test_data_records = get_oq_test_data_records()

    document.merge_rows('test_data_name', test_data_records)
    document.merge_rows('id_rep1', checklist_tables[0])
    document.merge_rows('id_rep2', checklist_tables[1])

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def prepare_system_specification():
    """Prepare the System Specification validation document
    :return None:
    """
    doc_type = 'System Specification'
    template_file = get_template_file(doc_type, SYSTEM_SPECIFICATION_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)
    
    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - System Specification - ' + g_document_prepared_date + '.docx'

    hardware_table_records = get_iq_hardware_table_records('IQ')
    software_table_records = get_iq_software_table_records('IQ')

    document.merge_rows('h_id', hardware_table_records)
    document.merge_rows('s_id', software_table_records)

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def prepare_test_plan():
    """Prepare the Test Plan validation document
    :return None:
    """

    doc_type = 'Test Plan'
    template_file = get_template_file(doc_type, TEST_PLAN_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)

    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - Test Plan - ' + g_document_prepared_date + '.docx'

    hardware_table_records = get_iq_hardware_table_records('IQ')
    software_table_records = get_iq_software_table_records('IQ')

    checklist_tables = get_oq_checklist_tables()

    test_data_records = get_oq_test_data_records()

    document.merge_rows('test_data_name', test_data_records)
    document.merge_rows('id_rep1', checklist_tables[0])

    document.merge_rows('h_id', hardware_table_records)
    document.merge_rows('s_id', software_table_records)

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def get_user_requirements_checklist_file(doc_type='User Requirements'):
    """Derive the User Requirements checklist file
    :param doc_type: {str} the document type, default 'User Requirements'
    :return infile: {str} the User Requirements checklist file
    """
    if doc_type not in g_config or 'checklist file basename' not in g_config[doc_type]:
        error_msg = "Could not retrieve the '{}' 'checklist file basename' from the config file".format(doc_type)
        logging.error(error_msg)
        raise Exception(error_msg)
    else:
        basename = g_config[doc_type]['checklist file basename']
        infile = os.path.dirname(os.path.abspath(__file__)) + '/conf/' + basename

    if not os.path.exists(infile):
        raise Exception("file '{}' does not exist".format(infile))

    return infile


def get_user_requirements_table_records(doc_type='User Requirements'):
    """Parse the User Requirements checklist tab-delimited file and build list of dictionaries
    :param doc_type: {str} document type, default 'User Requirements'
    :return hardware_table_records: {list} array of dictionaries
    """
    infile = get_user_requirements_checklist_file(doc_type)

    header_to_position_lookup = {}
    record_ctr = 0

    table_records = []
    id_ctr = 0

    with open(infile) as f:
        reader = csv.reader(f, delimiter='\t')
        row_ctr = 0
        for row in reader:
            row_ctr += 1
            if row_ctr == 1:
                field_ctr = 0
                for field in row:
                    header_to_position_lookup[field] = field_ctr
                    field_ctr += 1
                logging.info("Processed the header of csv file '{}'".format(infile))
            else:
                id_ctr += 1

                table_records.append({
                    'id': str(id_ctr),
                    'req': row[header_to_position_lookup['Requirement Description']],
                    'criticality': row[header_to_position_lookup['Criticality']],
                    'comment': ''
                })

                record_ctr += 1

        logging.info("Processed '{}' records in tab-delimited file '{}'".format(record_ctr, infile))

    logging.info(table_records)

    return table_records


def prepare_user_requirements():
    """Prepare the User Requirements validation document
    :return None:
    """
    doc_type = 'User Requirements'

    template_file = get_template_file(doc_type, USER_REQUIREMENTS_TEMPLATE_FILE_BASENAME)

    document = instantiate_mailmerge(template_file)

    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - User Requirements - ' + g_document_prepared_date + '.docx'

    user_req_table_records = get_user_requirements_table_records(doc_type)

    document.merge_rows('id', user_req_table_records)

    document.write(outfile)

    print("Wrote '{}' validation document  '{}'".format(doc_type, outfile))


def prepare_validation_report():
    """Prepare the Validation Report validation document
    :return None:
    """
    template_file = os.path.join(g_template_files_dir, VALIDATION_REPORT_TEMPLATE_FILE_BASENAME)
    outfile = g_outdir + '/' + g_software_name + ' ' + g_software_version + ' - Validation Report - ' + g_document_prepared_date + '.docx'
    prepare_validation_document(template_file, outfile)


@click.command()
@click.option('--verbose', is_flag=True, help='Will print more info to STDOUT')
@click.option('--outdir', help='The default is the current working directory')
@click.option('--config_file', type=click.Path(exists=True), help="The configuration file for this project")
@click.option('--logfile', help="The log file")
@click.option('--outfile', help="The output final report file")
@click.option('--template_files_dir', help="The directory containing the template files")
@click.option('--software_name', help="The name of the software system")
@click.option('--software_version', help="The version of the software system")
@click.option('--server', help="The server on which the software will be installed and validated on")
@click.option('--document_prepared_by', help="The name of the person that prepared the document")
@click.option('--document_prepared_date', help="The date the document was prepared")
def main(verbose, outdir, config_file, logfile, outfile, template_files_dir, software_name, software_version, server, document_prepared_by, document_prepared_date):
    """Template command-line executable
    """

    error_ctr = 0

    if software_name is None:
        print(Fore.RED + "--software_name was not specified")
        print(Style.RESET_ALL + '', end='')
        error_ctr += 1

    if software_version is None:
        print(Fore.RED + "--software_version was not specified")
        print(Style.RESET_ALL + '', end='')
        error_ctr += 1

    if server is None:
        print(Fore.RED + "--server was not specified")
        print(Style.RESET_ALL + '', end='')
        error_ctr += 1

    if error_ctr > 0:
        sys.exit(1)

    assert isinstance(software_name, str)
    assert isinstance(software_version, str)
    assert isinstance(server, str)

    if document_prepared_by is None:
        document_prepared_by = DEFAULT_document_prepared_by
        print(Fore.YELLOW + "--document_prepared_by was not specified and therefore was set to '{}'".format(document_prepared_by))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(document_prepared_by, str)

    if document_prepared_date is None:
        document_prepared_date = DEFAULT_DOCUMENT_PREPARED_DATE
        print(Fore.YELLOW + "--document_prepared_date was not specified and therefore was set to '{}'".format(document_prepared_date))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(document_prepared_date, str)
        
    if config_file is None:
        config_file = DEFAULT_CONFIG_FILE
        print(Fore.YELLOW + "--config_file was not specified and therefore was set to '{}'".format(config_file))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(config_file, str)

    if not os.path.exists(config_file):
        print(Fore.RED + "config_file '{}' does not exist".format(config_file))
        print(Style.RESET_ALL + '', end='')
        sys.exit(1)

    if outdir is None:
        outdir = DEFAULT_OUTDIR
        print(Fore.YELLOW + "--outdir was not specified and therefore was set to '{}'".format(outdir))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(outdir, str)

    if not os.path.exists(outdir):
        pathlib.Path(outdir).mkdir(parents=True, exist_ok=True)
        print(Fore.YELLOW + "Created output directory '{}'".format(outdir))
        print(Style.RESET_ALL + '', end='')

    if logfile is None:
        logfile = outdir + '/' + os.path.basename(__file__) + '.log'
        print(Fore.YELLOW + "--logfile was not specified and therefore was set to '{}'".format(logfile))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(logfile, str)

    if outfile is None:
        outfile = outdir + '/' + software_name + '_' + software_version + '_IQ_Checklist_' + document_prepared_date + '.docx'
        print(Fore.YELLOW + "--outfile was not specified and therefore was set to '{}'".format(outfile))
        print(Style.RESET_ALL + '', end='')

    assert isinstance(outfile, str)

    logging.basicConfig(filename=logfile, format=LOGGING_FORMAT, level=LOG_LEVEL)

    # Read the configuration from the JSON file and load into dictionary.
    logging.info("Loading configuration from '{}'".format(config_file))

    global g_config
    g_config = json.loads(open(config_file).read())

    if template_files_dir is None:
        if 'template_files_dir' not in g_config:
            template_files_dir = DEFAULT_TEMPLATE_FILES_DIR
            print(Fore.YELLOW + "--template_files_dir was not specified and therefore was set to default '{}'".format(template_files_dir))
            print(Style.RESET_ALL + '', end='')
        else:
            template_files_dir = config['template_files_dir']
            print(Fore.YELLOW + "--template_files_dir was not specified and therefore was set to '{}'".format(template_files_dir))
            print(Style.RESET_ALL + '', end='')

    assert isinstance(template_files_dir, str)

    if not os.path.exists(template_files_dir):
        print(Fore.RED + "template_files_dir '{}' does not exist".format(template_files_dir))
        print(Style.RESET_ALL + '', end='')
        sys.exit(1)

    global g_software_name
    global g_software_version
    global g_document_prepared_by
    global g_document_prepared_date
    global g_template_files_dir
    global g_outdir
    global g_server

    g_software_name = software_name
    g_software_version = software_version
    g_document_prepared_by = document_prepared_by
    g_document_prepared_date = document_prepared_date
    g_outdir = outdir
    g_template_files_dir = template_files_dir
    g_server = server

    prepare_iq()
    prepare_oq()
    prepare_pq()
    prepare_system_specification()
    prepare_test_plan()
    prepare_user_requirements()
    prepare_validation_report()


if __name__ == "__main__":
    main()