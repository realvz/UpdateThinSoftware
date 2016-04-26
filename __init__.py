from _winreg import *
import sys
import os
import subprocess
from shutil import copyfile
import logging
import smtplib
import shutil

view_client_guids = (r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{32F7C990-439B-4B2A-B472-A478D53D3F32}',
                     r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{44823D81-BDE3-4BAD-99B2-A7730F50818D}',
                     r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F0069EF8-0A18-4B53-8D18-697594146D59}',
                     r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B4264695-97BF-4E19-A760-62D8D531E5A1}')
imprivata_guids = (r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{E72D11F6-1F0C-47A9-B196-69BCF4132E66}',
                   r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{D386AC04-2AA1-44C2-A74E-9329B50D9495}',
                   r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{CED297D2-CB19-4544-AB28-259A8136A059}')

# ImprivataGUIDKey1 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{E72D11F6-1F0C-47A9-B196-69BCF4132E66}'
# ImprivataGUIDKey2 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{D386AC04-2AA1-44C2-A74E-9329B50D9495}'
# ImprivataGUIDKey3 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{CED297D2-CB19-4544-AB28-259A8136A059}'
# ViewClientGUIDKey1 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{32F7C990-439B-4B2A-B472-A478D53D3F32}'
# ViewClientGUIDKey2 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F0069EF8-0A18-4B53-8D18-697594146D59}'
# ViewClientGUIDKey3 = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{44823D81-BDE3-4BAD-99B2-A7730F50818D}'

DPGUIDKey = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{EA01643E-1D8A-43F1-8C23-32E9253CD08A}'
ImprivataInstallFile = 'c:\\temp\\ImprivataAgent.msi'
ViewClientInstallFile = 'c:\\temp\\VMware-Horizon-View-Client-x86-3.2.0-2331566.exe'
ViewClientInstallFile34 = 'c:\\temp\\VMware-Horizon-View-Client-x86-3.4.0-2769709.exe'
DefaultProfileData_W7 = 'c:\\temp\\NTUSER.DAT_W7'
DefaultProfileData_XP = 'c:\\temp\\NTUSER.DAT_XP'
LTAProfileData_W7 = 'c:\\temp\\NTUSER.DAT_W7_LTA'
LTAProfileData_XP = 'c:\\temp\\NTUSER.DAT_XP_LTA'
AppVersion = '0.9.3'
LogFilePath = 'C:\\TEMP\\ThinClientUpdater.log'

FoundImprivataGUID = ''


def log_global_vars():
    logging.info(
        '###########################################################################################################')
    logging.info('AppVersion=%s' % AppVersion)
    logging.info('Hostname=%s' % os.getenv('COMPUTERNAME', 'DummyHostname'))
    logging.info('ViewClientInstallFile34=%s' % ViewClientInstallFile34)
    logging.info('ImprivataInstallFile=%s' % ImprivataInstallFile)
    logging.info('ViewClientInstallFile=%s' % ViewClientInstallFile)
    logging.info(
        '###########################################################################################################')


def windows_version():
    import platform
    return platform.win32_ver()[0]


def copy_default_profile_file(filePath, destPath):
    if os.path.exists(filePath):
        if os.path.exists(destPath):
            os.remove(destPath)
        try:
            shutil.copyfile(filePath, destPath)
        except:
            print str(sys.exc_info()[0])
            print('Error copying NTUSER.DAT file.')
            return False
        return True
    else:
        print(u'{0:s} not found'.format(filePath))
        return False


def del_shell_vbs():
    filePath = 'c:\\windows\\shell.vbs'
    if os.path.exists(filePath):
        try:
            os.remove(filePath)
        except:
            print str(sys.exc_info()[0])
            print('Error deleting Shell.vbs.')
            return False
    return True


def deploy_default_dat_localthinadmin():
    if windows_version() == 'XP':
        print("OS is XP")
        if copy_default_profile_file(DefaultProfileData_XP, 'c:\\Documents And Settings\Default User\\NTUSER.DAT'):
            return True
    elif windows_version() == '7':
        print("OS is win7")
        if copy_default_profile_file(DefaultProfileData_W7, 'c:\\Users\Default User\\NTUSER.DAT'):
            return True
    else:
        print(u'Unknown Windows version found {0:s}'.format(windows_version()))


def remove_nla_proxy():
    nla_proxy_key = r'SYSTEM\CurrentControlSet\services\NlaSvc\Parameters\Internet\ManualProxies'
    try:
        reg_key = OpenKey(HKEY_LOCAL_MACHINE, nla_proxy_key, 0, KEY_ALL_ACCESS)
        SetValueEx( reg_key, "", 0, REG_SZ, "" )
    except:
        # We don't care about this step failing.
        print( 'Error disabling NLA proxy. {0}'.format(str(sys.exc_info()[0])) )
        logging.error( 'Error disabling NLA proxy. {0}'.format(str(sys.exc_info()[0])) )
    return True


def deploy_default_dat():
    if windows_version() == 'XP':
        print('OS is XP')
        if copy_default_profile_file(DefaultProfileData_XP, 'c:\\Documents And Settings\Default User\\NTUSER.DAT'):
            return True
    elif windows_version() == '7':
        print('OS is win7')
        if copy_default_profile_file(DefaultProfileData_W7, 'c:\\Users\Default User\\NTUSER.DAT'):
            return True
    else:
        print(u'Unknown Windows version found {0:s}'.format(windows_version()))


def upgrade_imprivata():
    print('Upgrading Imprivata...')
    if os.path.exists(ImprivataInstallFile):
        command = 'msiexec /i %s /qn /norestart' % ImprivataInstallFile
        print(command)
        logging.info(command)
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
        process.wait()

        if process.returncode == 3010:
            process.returncode = 0
        elif process.returncode == 1620:
            print('Install file is invalid')
    else:
        print("Imprivata install file missing.")
        logging.error('Imprivata install file not found. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return process.returncode


def view_upgrade_decision(view_version):
    if view_version == '5.2.1.937772':
        return True
    elif view_version == '2.3.3.18259':
        return True
    elif view_version == '3.0.0.19696':
        return True
    elif view_version == '3.2.0.24246':
        if windows_version() == 'XP':
            print('View Client is up to date. Version.')
            return False
        elif windows_version() == '7':
            return True
    elif view_version == '3.4.0.2769709':
        return False
    elif view_version == '0':
        return False
    else:
        print(u'View Client version unknown. {}'.format(view_version))
        return False


def upgrade_view_client():
    print('Upgrading View Client...')
    if windows_version() == 'XP':
        view_client_installer = ViewClientInstallFile
    else:
        view_client_installer = ViewClientInstallFile34

    # We are skipping View upgrade on all Windows 7 devices for now
    if windows_version() != 'XP':
        print('View Client install is being skipped on t520 thin clients')
        logging.info("View Client install is being skipped on t520 thin clients.")
        return 0

    if os.path.exists(view_client_installer):
        print('%s  /S /V /qn ADDLOCAL=ALL REBOOT=Reallysuppresss' % view_client_installer)
        logging.info('%s  /S /V /qn ADDLOCAL=ALL REBOOT=Reallysuppresss' % view_client_installer)
        command = '%s  /S /V /qn ADDLOCAL=ALL REBOOT=Reallysuppresss' % view_client_installer
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
        process.wait()

        if process.returncode == 3010:
            process.returncode = 0
        elif process.returncode == 1620:
            print('Install file is invalid')
            logging.error('Install file is invalid.')
    else:
        print('View Client install file missing.')
        logging.error('Could not locate View Client install file. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return process.returncode


def hide_view_shade():
    view_client_registry_key = r'Software\VMware, Inc.\VMware VDM\Client'
    try:
        reg_key = OpenKey(HKEY_LOCAL_MACHINE, view_client_registry_key, 0, KEY_ALL_ACCESS)
        SetValueEx(reg_key, "EnableShade", 0, REG_SZ, 'false')
    except:
        print('error disabling view shade.')
        logging.error('Could not disable View shade. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return 0


def register_vmusb_sys():
    vmusb_sys_path = 'C:\\Program Files\\Common Files\\VMware\\USB\\vmusb.sys'
    vmusb_inf_path = 'C:\\Program Files\\Common Files\\VMware\\USB\\vmusb.inf'
    if os.path.exists(vmusb_sys_path):
        try:
            copyfile(vmusb_sys_path, 'C:\\windows\system32\\drivers\\vmusb.sys')
        except:
            print('Could not copy vmusb.sys file to system32')
            logging.error('Could not copy vmusb.sys to system32. ' + str(sys.exc_info()[0]))
    command = 'rundll32.exe advpack.dll,LaunchINFSectionEx \"%s\",,,4' % vmusb_inf_path
    process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    process.wait()


def get_view_client_version(view_client_guid):
    try:
        reg_key = OpenKey(HKEY_LOCAL_MACHINE, view_client_guid, 0, KEY_ALL_ACCESS)
        view_client_version = QueryValueEx(reg_key, "DisplayVersion")[0]
    except:
        logging.error('Could not get View Client Version. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return view_client_version


def update_imprivata_appliance_addr():
    imprivata_reg_key = 'SOFTWARE\SSOProvider\ISXAgent'
    try:
        reg_key = OpenKey(HKEY_LOCAL_MACHINE, imprivata_reg_key, 0, KEY_ALL_ACCESS)
        SetValueEx(reg_key, "IPTXPrimServer", 0, REG_SZ,
                   'https://QorImprivata.gtm.iqor.qor.com/sso/servlet/messagerouter')
    except:
        print('error setting IPTXPrimServer value.')
        logging.error('error setting IPTXPrimServer value. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return 0


def get_imprivata_version(impGUID):
    try:
        reg_key = OpenKey(HKEY_LOCAL_MACHINE, impGUID, 0, KEY_ALL_ACCESS)
        imprivata_version = QueryValueEx(reg_key, "DisplayVersion")[0]
    except:
        print('error getting imprivata version')
        logging.error('Could not get Imprivata Agent Version. {0}'.format(str(sys.exc_info()[0])))
        return -1
    return imprivata_version


def is_thin_imprivata(impGUID):
    global FoundImprivataGUID
    try:
        OpenKey(HKEY_LOCAL_MACHINE, impGUID, 0, KEY_ALL_ACCESS)
    except WindowsError:
        logging.error(
            u'Could not determine if thin client is thin is Imprivata. Error reading registry. {0:s} '.format(impGUID))
        return False
    except:
        e = sys.exc_info()[0]
        print e
        logging.error('Could not determine if thin client is thin is Imprivata. {0}'.format(str(sys.exc_info()[0])))
        return False
    FoundImprivataGUID = impGUID
    return True


def is_thin_DP():
    try:
        OpenKey(HKEY_LOCAL_MACHINE, DPGUIDKey, 0, KEY_READ)
        return True
    except:
        return False


def send_email():
    from email.mime.text import MIMEText
    host_name = os.getenv('COMPUTERNAME', 'DummyHostname')
    fp = open(LogFilePath, 'rb')
    msg = MIMEText(fp.read())
    fp.close()
    msg['Subject'] = 'Thin upgrade task failed'
    msg['From'] = '%s@iqor.com' % host_name
    # msg['To'] = 're.alvarez@iqor.com'
    s = smtplib.SMTP('exchange-relay.iqor.qor.com')
    s.sendmail('h9489893842@iqor.com', ['re.alvarez@iqor.com', 'john.schulze@iqor.com','qdi$imaging@iqor.com'], msg.as_string())


def main():
    perform_view_upgrade = False
    did_anything_fail = False
    view_client_version = 0
    imprivata_version = 0

    if not os.path.exists('c:\\temp\\'):
        os.mkdir('c:\\temp')

    logging.basicConfig(level=logging.DEBUG, filename=LogFilePath,
                        filemode="a+", format="%(asctime)-15s %(levelname)-8s %(message)s")

    log_global_vars()

    print('Initializing thin client upgrade script...')
    logging.info('Initializing thin client upgrade script...')

    if is_thin_DP():
        logging.info('This is a DP thin client.')
    else:
        for i in range(0, len(imprivata_guids)):
            imprivata_version = get_imprivata_version(imprivata_guids[i])
            if imprivata_version <> -1:
                print('Imprivata version is %s' % imprivata_version)
                logging.info('Imprivata version is %s' % imprivata_version)
                if imprivata_version in ('4.9.110.68', '4.9.199.1087'):
                    if upgrade_imprivata() == 0:
                        print('Imprivata agent upgraded successfully.')
                        logging.info('Imprivata agent upgraded successfully.')
                    else:
                        print('Fail flag tripped')
                        did_anything_fail = True
                        print('Imprivata agent upgrade failed.')
                        logging.info('Imprivata agent upgraded failed.')
                        break
                elif imprivata_version == '5.1.003.22':
                    print('Imprivata agent is up to date.')
                    logging.info('Imprivata agent is up to date.')
                    break

        if imprivata_version == 0:
            print('This is not an imprivata thin client.')
            logging.info('This is not an imprivata thin client.')

    for i in range(0, len(view_client_guids)):
        view_guid = view_client_guids[i]
        view_client_version = get_view_client_version(view_guid)
        if view_client_version <> -1:
            print(u'Current View client version is {0:s}'.format(view_client_version))
            logging.info(u'Current View client version is {0:s}'.format(view_client_version))
            break

    if view_upgrade_decision(view_client_version):
        if upgrade_view_client() == 0:
            register_vmusb_sys()
            logging.info('View client upgraded successfully.')
            print('View client upgraded successfully.')
        else:
            print('Fail flag tripped')
            did_anything_fail = True
            logging.info('View client upgrade failed.')
            print('View client upgrade failed.')

    if deploy_default_dat():
        logging.info('Default User profile deployed')
        print('Default User profile deployed')
    else:
        logging.info('Default User profile deployment failed.')
        print('Default User profile deployment failed.')
        did_anything_fail = True

    if deploy_default_dat_localthinadmin():
        logging.info('Localthinadmin User profile deployed')
        print('Localthinadmin User profile deployed')
    else:
        logging.info('Localthinadmin User profile deployment failed.')
        print('Localthinadmin User profile deployment failed.')
        did_anything_fail = True

    if hide_view_shade() <> 0:
        did_anything_fail = True

    if not is_thin_DP():
        if update_imprivata_appliance_addr() <> 0:
            did_anything_fail = True

    if not del_shell_vbs():
        did_anything_fail = True

    if not remove_nla_proxy():
        did_anything_fail = True

    if did_anything_fail:
        # Send an email with the log file
        send_email()


if __name__ == '__main__':
    main()
