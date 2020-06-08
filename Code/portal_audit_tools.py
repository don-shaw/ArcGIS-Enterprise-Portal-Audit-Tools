from arcgis.gis import GIS, RoleManager
from datetime import datetime
import pandas as pd
import csv
from os import path
import os
import configparser
import logging
import keyring
import subprocess
import warnings
import arcpy
import time
warnings.filterwarnings('ignore')


def create_directories(report_dir, today_dir):
    os.chdir(report_dir)
    if path.isdir(today_dir) is False:
        logging.info(f'Creating  {today_dir}...')
        os.mkdir(path.join(report_dir, today_dir))
        os.chdir(today_dir)
        os.mkdir('csv_files')


def generate_sys_log_report(slp_directory, today_dir, server_log_dir):
    if path.isdir(today_dir) is True:
        os.chdir(today_dir)
        os.mkdir('sys_log_report')
        os.chdir(slp_directory)
        output_dir = path.join(today_dir, 'sys_log_report')
        cmd = (str(f"slp.exe -f AGSFS -i {server_log_dir} -d {output_dir} -eh now -sh 1080 -a complete -r spreadsheet -sbu true -o false, shell=True"))
        logging.info('Creating the System Log Report...')
        subprocess.call(cmd)



def get_portal_data(url, cred_name, portal_username, today_dir):

    try:
        # Get credentials
        portal_password = keyring.get_password(cred_name, portal_username)
        portal = GIS(url, portal_username, portal_password, verify_cert=False)
        logging.info('Querying the Enterprise Portal...')

        # Query users, groups, and items
        users = portal.users.search('!username:esri_*')
        groups = portal.groups.search('!owner:esri_*')
        all_items = portal.content.search(query='!owner:esri*', max_items=10000)

        # Create empty dictionaries, will be used to populate CSV files
        user_dict = {}
        group_dict = {}
        item_dict = {}

        # Get users and write them to CSV
        with open(path.join(today_dir, 'csv_files', 'users.csv'), 'w', newline='', encoding='utf-8') as user_csv:
            user_file = csv.DictWriter(user_csv,
                                       fieldnames=['USERNAME', 'EMAIL', 'ROLE', 'LAST_LOGIN',
                                                   'CREATED', 'GROUPS', 'ITEMS'])
            user_file.writeheader()

            rm = RoleManager(portal)
            roles = rm.all()

            for user in users:
                # if user.idpUsername is not None:
                num_items = 0

                user_dict['USERNAME'] = user.username
                user_dict['EMAIL'] = user.email
                user_dict['ROLE'] = user.role
                for role in roles:
                    role_id = role.role_id
                    if role_id == user.roleId:
                        user_dict['ROLE'] = role.name

                if user.lastLogin != -1:
                    user_dict['LAST_LOGIN'] = datetime.fromtimestamp(float(user.lastLogin / 1000)).strftime('%m/%d/%Y')
                else:
                    user_dict['LAST_LOGIN'] = -1
                user_dict['CREATED'] = datetime.fromtimestamp(float(user.created / 1000)).strftime('%m/%d/%Y')

                user_groups = user.groups
                g_list = []

                for g in user_groups:
                    g_list.append(g.title)
                user_dict['GROUPS'] = str(g_list)[1:-1]

                user_content = user.items()
                folders = user.folders

                for item in user_content:
                    num_items += 1

                for folder in folders:
                    folder_items = user.items(folder=folder['title'])
                    for item in folder_items:
                        num_items += 1
                user_dict['ITEMS'] = num_items
                user_file.writerow(user_dict)
        logging.info('User File:    {0}'.format(path.join(today_dir, 'csv_files', 'user.csv')))

        # Get Groups
        with open(path.join(today_dir, 'csv_files', 'groups.csv'), 'w', newline='', encoding='utf-8') as group_csv:
            groups_file = csv.DictWriter(group_csv, fieldnames=['TITLE', 'OWNER', 'MANAGERS', 'USERS', 'ITEMS'])
            groups_file.writeheader()

            for g in groups:

                members = g.get_members()

                group_dict['TITLE'] = g.title
                group_dict['OWNER'] = members['owner']
                group_dict['MANAGERS'] = str(str(members['admins']).replace("'", ''))[1:-1]
                group_dict['USERS'] = str(str(members['users']).replace("'", ''))[1:-1]
                group_dict['ITEMS'] = len(g.content())
                groups_file.writerow(group_dict)
        logging.info('Group File:    {0}'.format(path.join(today_dir, 'csv_files', 'groups.csv')))

        # Get Items
        with open(path.join(today_dir, 'csv_files', 'items.csv'), 'w', newline='', encoding='utf-8') as items_csv:
            items_file = csv.DictWriter(items_csv, fieldnames=['TITLE', 'OWNER', 'ID', 'TYPE', 'AUTHORITATIVE',
                                                               'TAGS', 'ACCESS', 'SHARED_WITH_ORG',
                                                               'SHARED_WITH_EVERYONE', 'SHARED_WITH_GROUPS', 'VIEWS',
                                                               'CREATED', 'HOMEPAGE', 'THUMBNAIL', 'DESCRIPTION', 'SIZE'])

            items_file.writeheader()
            for item in all_items:
                if item.type in ['Geoprocessing Service', 'Service Definition', 'Code Attachment', 'Geometry Service',
                                 'Vector Tile Service', 'Vector Tile Package']:
                    pass
                else:
                    print(item)
                    item_groups = []
                    item_dict['TITLE'] = item.title
                    item_dict['OWNER'] = item.owner
                    item_dict['ID'] = item.id
                    item_dict['TYPE'] = item.type
                    item_dict['AUTHORITATIVE'] = item.content_status
                    item_dict['TAGS'] = str(item.tags)[1:-1]
                    item_dict['DESCRIPTION'] = item.description
                    item_dict['VIEWS'] = item.numViews
                    item_dict['CREATED'] = datetime.fromtimestamp(float(item.created / 1000)).strftime('%m/%d/%Y')
                    # item_dict['dependent_on'] = item.dependent_upon()
                    # item_dict['dependent_to'] = item.dependent_to()
                    item_dict['HOMEPAGE'] = item.homepage
                    item_dict['SHARED_WITH_EVERYONE'] = item.shared_with['everyone']
                    item_dict['SHARED_WITH_ORG'] = item.shared_with['org']
                    for g in item.shared_with['groups']:
                        item_groups.append(g.title)
                    item_dict['SHARED_WITH_GROUPS'] = str(item_groups)[1:-1]
                    # print(item_groups)
                    item_dict['ACCESS'] = item.access
                    item_dict['SIZE'] = item.size / 1000 / 1000
                    item_dict['THUMBNAIL'] = item.thumbnail
                    items_file.writerow(item_dict)

        logging.info('Item File:    {0}'.format(path.join(today_dir, 'csv_files', 'items.csv')))
        return groups
    except Exception as e:
        logging.error(e)


def process_sys_log_report(today_dir):

    try:
        report_dir = path.join(today_dir, 'sys_log_report')
        os.chdir(report_dir)
        for file in os.listdir(report_dir):
            if file.endswith('xlsx'):
                report = file
        # System Log Parser dfs
        stats_by_user = pd.read_excel(report, sheet_name='Statistics By User', header=4)
        stats_by_resource = pd.read_excel(report, sheet_name='Statistics By Resource', header=4)
        resources_by_time = pd.read_excel(report, sheet_name='Elapsed Time - All Resources', header=3)

        # Items DF
        items_df = pd.read_csv(path.join(today_dir, 'csv_files', 'items.csv'))

        # Throughput
        throughput = pd.read_excel(report, sheet_name='Throughput per Minute', header=3)
        throughput['date'] = pd.to_datetime(throughput['Date Time (Local Time)']).dt.to_period('D')
        throughput['Date Time (Local Time)'] = pd.to_datetime(throughput['Date Time (Local Time)'])
        throughput['Date Time (Local Time)'].dt.strftime('%m/%d/%Y %H:%M:%S')
        throughput.rename(columns={'Date Time (Local Time)': 'Date_Time', 'Epoch Time': 'Epoch_Time',
                                   'Requests/Minute': 'Requests_Minute', 'Requests/Seccond': 'Requests_Seccond',
                                   'Avg Response Time': 'Avg_Response_Time', 'Min Response Time': 'Min_Response_Time',
                                   'P95 Response Time': 'P95_Response_Time', 'P99 Response Time': 'P99_Response_Time',
                                   'Max Response Time': 'Max_Response_Time', 'HTTP 200': 'HTTP_200', 'HTTP 300': 'HTTP_300',
                                   'HTTP 400': 'HTTP_400', 'HTTP 500': 'HTTP_500'}, inplace=True)
        throughput.to_csv(path.join(today_dir, 'csv_files', 'throughput.csv'), index=False)
        logging.info('Throughput File:    {0}'.format(path.join(today_dir, 'csv_files', 'throughput.csv')))

        # Stats by user
        stats_by_user = stats_by_user[stats_by_user['Resource'].str.contains('GPServer') == False]
        stats_by_user = stats_by_user[stats_by_user['User'] != '-']
        stats_by_user.rename(columns={'Count Pct': 'Count_Pct', 'Sum Pct': 'Sum_Pct'}, inplace=True)
        stats_by_user.to_csv(path.join(today_dir, 'csv_files', 'stats_by_user.csv'), index=False)
        logging.info('Throughput File:    {0}'.format(path.join(today_dir, 'csv_files', 'stats_by_user.csv')))


        # Stats by resource
        stats_by_resource.groupby('Resource')['Count'].sum().reset_index()
        stats_by_resource = stats_by_resource[
            stats_by_resource['Resource'].isin(['All Resources',
                                                'Hover over each column header for description']) == False]
        stats_by_resource = stats_by_resource[stats_by_resource['Capability'].isin(['GPServer']) == False]
        stats_by_resource.rename(columns={'Count Pct': 'Count_Pct', 'Sum Pct': 'Sum_Pct'}, inplace=True)

        stats_by_resource.to_csv(path.join(today_dir, 'csv_files', 'stats_by_resource.csv'), index=False)
        logging.info('Stats by Resource File:    {0}'.format(path.join(today_dir, 'csv_files', 'stats_by_resource.csv')))


        # Item_Metrics
        items_df = items_df[items_df['TYPE'].isin(['Map Service', 'Feature Service'])]
        items_df.loc[items_df['TYPE'] == 'Map Service', 'Resource'] = items_df['TITLE'] + '.MapServer'
        items_df.loc[items_df['TYPE'] == 'Feature Service', 'Resource'] = items_df['TITLE'] + '.FeatureServer'
        last_accessed = resources_by_time.groupby('Resource')['Date Time (Local Time)'].max().reset_index()
        last_accessed['LAST_ACCESSED'] = pd.to_datetime(last_accessed['Date Time (Local Time)']).dt.to_period('D')
        last_accessed = last_accessed[last_accessed['Resource'].str.contains('GPServer') == False]
        last_accessed.loc[last_accessed.Resource.str.contains('/'), 'Resource'] = last_accessed.Resource.str.split("/", expand=True)[1]
        last_accessed.rename(columns={'Date Time (Local Time)': 'Date_Time'}, inplace=True)
        item_metrics = items_df.merge(right=last_accessed, how='outer', left_on='Resource', right_on='Resource')

        item_metrics.to_csv(path.join(today_dir, 'csv_files', 'item_metrics.csv'), index=False)
        logging.info('Item Metrics File:    {0}'.format(path.join(today_dir, 'csv_files', 'item_metrics.csv')))

    except Exception as processing_error:
        logging.error(processing_error)

def process_fgdb(fgdb, today_dir, items, groups, users, item_metrics, stats_by_resource, stats_by_user, throughput):

    users_csv = path.join(today_dir, 'csv_files', 'users.csv')
    items_csv = path.join(today_dir, 'csv_files', 'items.csv')
    groups_csv = path.join(today_dir, 'csv_files', 'groups.csv')
    throughput_csv = path.join(today_dir, 'csv_files', 'throughput.csv')
    item_metrics_csv = path.join(today_dir, 'csv_files', 'item_metrics.csv')
    stats_by_user_csv = path.join(today_dir, 'csv_files', 'stats_by_user.csv')
    stats_by_resource_csv = path.join(today_dir, 'csv_files', 'stats_by_resource.csv')

    try:
        logging.info('Processing fgdb...')

        #logging.info(fgdb)
        arcpy.management.TruncateTable(users)
        arcpy.management.TruncateTable(groups)
        arcpy.management.TruncateTable(items)
        arcpy.management.TruncateTable(throughput)
        arcpy.management.TruncateTable(stats_by_resource)
        arcpy.management.TruncateTable(stats_by_user)
        arcpy.management.TruncateTable(item_metrics)

        arcpy.Append_management(users_csv, users, "TEST")
        arcpy.Append_management(groups_csv, groups, "TEST")
        arcpy.Append_management(items_csv, items, "TEST")
        arcpy.Append_management(throughput_csv, throughput, "TEST")
        arcpy.Append_management(item_metrics_csv, item_metrics, "TEST")
        arcpy.Append_management(stats_by_resource_csv, stats_by_resource, "TEST")
        arcpy.Append_management(stats_by_user_csv, stats_by_user, "TEST")

    except Exception as error:
        logging.error(error)


if __name__ == '__main__':

    log_dir = path.dirname(path.realpath(__file__))
    log_name = path.join(log_dir, 'portal_audit_tools.log')
    log_format = "%(asctime)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename=log_name,
                        level=logging.INFO,
                        format=log_format,
                        filemode="w")
    logging.getLogger('arcgis').setLevel(logging.WARNING)

    config = configparser.ConfigParser()
    config.read(path.join(log_dir, 'config.ini'))

    portal_url = config.get('ALL', 'portal_url')
    portal_cred_name = config.get('ALL', 'portal_cred_name')
    portal_cred_user = config.get('ALL', 'portal_cred_user')
    reports_directory = config.get('ALL', 'reports_directory')
    #today_directory = path.join(reports_directory, f'{datetime.now().strftime("%m-%d-%Y")}')
    today_directory = r"C:\Data\Reports\06-05-2020"
    system_log_parser = config.get('ALL', 'sys_log_directory')
    server_log_directory = config.get('ALL', 'server_log_directory')
    file_geodatabase = config.get('ALL', 'file_geodatabase')
    users_table =  config.get('ALL', 'users_table')
    groups_table =  config.get('ALL', 'groups_table')
    items_table =  config.get('ALL', 'items_table')
    item_metrics_table =  config.get('ALL', 'item_metrics_table')
    stats_by_user_table = config.get('ALL', 'stats_by_user_table')
    stats_by_resource_table =  config.get('ALL', 'stats_by_resource_table')
    throughput_table =  config.get('ALL', 'throughput_table')


    logging.info("***** Start time:  {0}\n".format(datetime.now().strftime("%A %B %d %I:%M:%S %p %Y")))
    logging.info('Portal URL:   {0}'.format(portal_url))
    logging.info('Windows Credential:   {0}'.format(portal_cred_name))
    logging.info('Portal Username:  {0}'.format(portal_cred_user))
    logging.info('Audit Report Directory:   {0}\n'.format(today_directory))
    logging.info('FGDB:   {0}'.format(file_geodatabase))

    # Go!
    try:
        create_directories(reports_directory, today_directory)
        generate_sys_log_report(system_log_parser, today_directory, server_log_directory)
        get_portal_data(portal_url , portal_cred_name, portal_cred_user, today_directory)
        process_sys_log_report(today_directory)
        process_fgdb(file_geodatabase, today_directory, items_table, groups_table, users_table,
                     item_metrics_table, stats_by_resource_table, stats_by_user_table, throughput_table)
    except Exception as e:
        print(e)
    logging.info("***** Completed time:  {0}\n".format(datetime.now().strftime("%A %B %d %I:%M:%S %p %Y")))
    logging.shutdown()
