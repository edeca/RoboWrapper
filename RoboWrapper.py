import win32com.client
import os.path
import subprocess
import sys
import yaml
import pytz
import argparse
import errno
import glob
import logging
from datetime import datetime
from prettytable import PrettyTable

########
# Author: David <david@edeca.net>
#   Date: March 2014
#
# This script is a simple wrapper around Robocopy, intended to assist
# with drive letters that move around (e.g. USB drives, CD, network
# shares).
#
# Jobs are defined as simple YAML documents which describe the source
# and destination, Robocopy options and an optional log.
#
# See README.HTML for installation, configuration and usage details.
########

VERSION_MAJOR = 0
VERSION_MINOR = 1
BANNER = """
 ____       _        __        __
|  _ \ ___ | |__   __\ \      / / __ __ _ _ __  _ __   ___ _ __
| |_) / _ \| '_ \ / _ \ \ /\ / / '__/ _` | '_ \| '_ \ / _ \ '__|
|  _ < (_) | |_) | (_) \ V  V /| | | (_| | |_) | |_) |  __/ |
|_| \_\___/|_.__/ \___/ \_/\_/ |_|  \__,_| .__/| .__/ \___|_|
             by David Cannings (@edeca)  |_|   |_|     v{}.{}
"""

# Environment variables which will be expanded
VALID_ENV_VARS = ['SYSTEMROOT', 'TMP', 'COMPUTERNAME', 'USERDOMAIN',
        'PROGRAMFILES', 'PROGRAMFILES(X86)', 'COMMONPROGRAMFILES(X86)',
        'ALLUSERSPROFILE', 'LOCALAPPDATA', 'HOMEPATH', 'PROGRAMW6432',
        'USERNAME', 'PROGRAMDATA', 'WINDIR', 'APPDATA', 'HOMEDRIVE',
        'SYSTEMDRIVE', 'COMMONPROGRAMW6432', 'PUBLIC', 'USERPROFILE']

# Cache first WMI call as it can be slow
wmiLogicalDisks = None


class RoboException(Exception):
    pass


def GetLogicalDrivesFromWMI():
    global wmiLogicalDisks

    # Speedup: underlying WMI calls only made once then cached
    if wmiLogicalDisks is None:
        strComputer = "."
        objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        objSWbemServices = objWMIService.ConnectServer(strComputer,
            "root\cimv2")
        wmiLogicalDisks = objSWbemServices.ExecQuery(
            "Select * from Win32_LogicalDisk")

    return wmiLogicalDisks


def FindDriveFromSerial(serial):
    """ Find drive letter from a serial number (e.g. 'ABCD1234') """

    for objItem in GetLogicalDrivesFromWMI():
        if objItem.VolumeSerialNumber == serial:
            return objItem.Name

    return None


def FindDriveFromName(name):
    """ Find drive letter from volume name (e.g. 'KINGSTON') """

    for objItem in GetLogicalDrivesFromWMI():
        if objItem.VolumeName == name:
            return objItem.Name

    return None


def ResolveDriveLetter(options):
    """ Find drive letter from source or destination settings """

    tried = []

    # The first likely match is returned
    if 'serial' in options:
        tried.append("serial: {}".format(options['serial']))

        found = FindDriveFromSerial(options['serial'])
        if found:
            return found

    if 'name' in options:
        tried.append("name: {}".format(options['name']))

        found = FindDriveFromName(options['name'])
        if found:
            return found

    raise RoboException("Couldn't find drive letter, tried: {}".format(
        ", ".join(tried)))


def ListDrives():
    """ Print a table of currently available drives """

    table = PrettyTable(["Drive", "Serial", "Name"])
    for objItem in GetLogicalDrivesFromWMI():
        table.add_row([objItem.Name, objItem.VolumeSerialNumber,
                objItem.VolumeName])

    print table


def CheckFlag(options, path):
    """ Check the 'safety' flag and ensure it exists """

    # Flag is optional, return success if not specified
    if 'flag' not in options:
        return True

    flag = options['flag']
    flag = flag.replace('$drive$', os.path.splitdrive(path)[0])
    # It makes little sense to use $path$ with Robocopy's /MIR option,
    # which deletes files.  The current log won't be deleted (because
    # Robocopy will have the file open) but older logs will be removed.
    flag = flag.replace('$path$', path)

    logging.debug("Checking flag: {} ".format(flag))

    return os.path.isfile(flag)


def DoRobocopy(job, dry_run):
    """ Call robocopy and parse exit status for success """

    cmd = []
    cmd.append('robocopy')
    cmd.append(job['run']['src_path'])
    cmd.append(job['run']['dst_path'])
    cmd.append(job['run']['file_types'])
    cmd.extend(job['run']['options'])

    logging.debug("Command: {}".format(" ".join(cmd)))

    if dry_run:
        return True

    # TODO: If debug verbosity is set then don't discard output
    DEVNULL = open(os.devnull, 'wb')
    proc = subprocess.Popen(cmd, stdout=DEVNULL, stderr=DEVNULL)

    logging.debug("Robocopy running (PID: {})".format(proc.pid))

    # TODO: Allow user to provide an optional timeout and kill?
    proc.wait()

    logging.debug("Robocopy finished, return code: {}".format(proc.returncode))

    # TODO: Parse status code fully (see http://ss64.com/nt/robocopy-exit.html)
    if proc.returncode >= 8:
        return False

    return True


def ValidateJob(job):
    """ Validate loaded job for required settings """

    if 'name' not in job:
        raise RoboException("A name is required")

    if 'source' not in job:
        raise RoboException("Source settings missing!")

    if 'destination' not in job:
        raise RoboException("Destination settings missing!")

    if 'path' not in job['source']:
        raise RoboException("Source path not defined")

    if 'path' not in job['destination']:
        raise RoboException("Destination path not defined")

    # TODO: Check serial, should be 8 hex characters

    return True


def ExpandEnvironmentVars(path):
    """ Expand environment variables """
    for var in VALID_ENV_VARS:
        try:
            path = path.replace("${}$".format(var), os.environ[var])
        except KeyError:
            pass

    return path


def SubstitutePath(path, job):
    """ Substitute placeholders with derived values """

    path = path.replace("$src_path$", job['run']['src_path'])
    path = path.replace("$dst_path$", job['run']['dst_path'])
    path = path.replace("$src_drive$", job['run']['src_drive'])
    path = path.replace("$dst_drive$", job['run']['dst_drive'])

    # Format timestamp
    timestamp = job['run']['time'].strftime(job['run']['time_format'])
    path = path.replace("$timestamp$", timestamp)

    path = ExpandEnvironmentVars(path)

    return path


def ParseSettings(job):
    """ Parse generic job options """

    # Optional setting: Extract the timestamp format (used for logs)
    try:
        job['run']['time_format'] = job['settings']['time_format']
    except KeyError:
        pass

    return job


def ParseRobocopyOptions(job):
    """ Parse Robocopy options (all optional) """

    # Optional setting: Robocopy settings
    try:
        job['run']['options'] = job['robocopy']['options'].split(' ')
    except KeyError:
        pass

    # Optional setting: Output file for Robocopy logs
    try:
        log_path = SubstitutePath(job['robocopy']['log'], job)
        logging.debug("Log will be saved to: {}".format(log_path))
        job['run']['options'].append("/LOG:{}".format(log_path))
    except KeyError:
        pass

    # Optional setting: Robocopy files to copy
    try:
        # Parse out the files option, should be a glob of extensions like "*.*"
        job['run']['file_types'] = job['robocopy']['files']
    except KeyError:
        pass

    return job


def ParseLocations(job):
    """ Find source and destination paths for a job """

    # Strategy is: find drive letter, expand path, check it exists

    job['run']['src_path'] = ExpandEnvironmentVars(job['source']['path'])
    job['run']['dst_path'] = ExpandEnvironmentVars(job['destination']['path'])

    # These keys trigger drive letter lookup
    k = ['serial', 'name']
    if any(name in k for name in job['source'].keys()):
        job['run']['src_drive'] = ResolveDriveLetter(job['source'])
        job['run']['src_path'] = os.path.join(job['run']['src_drive'],
                '\\', job['run']['src_path'])

    # Otherwise we need to derive the drive letter from the path
    else:
        job['run']['src_drive'] = os.path.splitdrive(job['run']['src_path'])[0]

    # TODO: Remove duplicated code.  This will require refactoring job['run']
    #       into source and destination sections
    if any(name in k for name in job['destination'].keys()):
        job['run']['dst_drive'] = ResolveDriveLetter(job['destination'])
        job['run']['dst_path'] = os.path.join(job['run']['dst_drive'],
                '\\', job['run']['dst_path'])

    # Otherwise we need to derive the drive letter from the path
    else:
        job['run']['dst_drive'] = os.path.splitdrive(job['run']['dst_path'])[0]

    # Check source path exists.  Destination path may not exist yet but can
    # be created by Robocopy.
    if not os.path.exists(job['run']['src_path']):
        raise RoboException("Source doesn't exist: {}".format(
                job['run']['src_path']))

    return job


def DefaultJobSettings(job):
    """ Setup default options which can be replaced with per-job settings """

    job['run'] = dict()

    # TODO: Read default options from somewhere to allow setting of
    #       these globally for all jobs, without specifying in every
    #       job file
    job['run']['options'] = []
    job['run']['file_types'] = "*.*"
    # Default is similar to UTC format without colons (invalid on Windows)
    job['run']['time_format'] = "%Y-%m-%dT%H%M%S%z"
    job['run']['time'] = pytz.utc.localize(datetime.now())

    return job


def LoadJob(fp):
    """ Load YAML from a stream and handle any errors """

    try:
        job = yaml.load(fp, Loader=yaml.SafeLoader)
    except yaml.YAMLError, exc:
        logging.error("YAML error in job file {} (run with -v for debug)"
                .format(fp.name))
        logging.debug("YAML parser output: {}".format(exc))
        return None

    return job


def RunJob(job_file, dry_run=False):
    # Check job is a file
    if not os.path.isfile(job_file):
        logging.error("Could not find file {}".format(job_file))
        return False

    # Load and parse YAML data
    try:
        fp = open(job_file, 'r')
    except IOError as e:
        # Check for permissions problems in a safe way
        if e.errno == errno.EACCES:
            logging.error("Could not access file {} (permission denied)"
                    .format(job_file))
            return False
        raise
    else:
        with fp:
            job = LoadJob(fp)
            if job is None:
                return False

    try:
        ValidateJob(job)
        logging.debug("Loaded job: {}".format(job['name']))
    except RoboException, exc:
        logging.error("Configuration problem: {} (in {})"
                .format(exc, job_file))
        return False

    try:
        job = DefaultJobSettings(job)
        job = ParseSettings(job)
        job = ParseLocations(job)
        job = ParseRobocopyOptions(job)

    except RoboException, exc:
        logging.error("{}".format(exc))
        return False

    # Check for safety flags in drives
    # TODO: Move into a generic JobReady() function
    if not CheckFlag(job['source'], job['run']['src_drive']):
        logging.error("Could not find source safety flag")
        return False

    if not CheckFlag(job['destination'], job['run']['dst_drive']):
        logging.error("Could not find destination safety flag")
        return False

    return DoRobocopy(job, dry_run)


def main():
    succeeded = 0
    failed = 0

    logging.basicConfig(format='%(levelname)8s: %(message)s')

    parser = argparse.ArgumentParser(description=
            'Run robocopy in a world of changing drive letters.')
    parser.add_argument('job', nargs='*',
        help="job files to run (can be a glob like *.yaml)")
    parser.add_argument('-d', '--dry-run', dest='dry_run', action='store_true',
        default=False, help="test jobs but don't run robocopy")
    parser.add_argument('-v', '--verbose', dest="verbosity", action='count',
        default=0, help="enable debug output")
    parser.add_argument('--drives', dest='drives', action='store_true',
        default=None, help="list information about available drives")

    args = parser.parse_args()

    print BANNER.format(VERSION_MAJOR, VERSION_MINOR)

    # Each -v option removes 10 from the default level of INFO
    level = 20 - args.verbosity * 10
    logging.getLogger('').setLevel(level)

    if args.drives:
        ListDrives()
        sys.exit(0)

    if not len(args.job):
        logging.error("Need at least one job (or --drives), see --help")
        sys.exit(1)

    # Don't allow multiple file globs to find the same job file twice
    jobs = set()

    for path in args.job:
        for file in glob.glob(path):
            jobs.add(file)

    logging.info("Found {} job(s) to run".format(len(jobs)))

    if args.dry_run:
        logging.warning("--dry-run option given, won't execute Robocopy")

    for job in jobs:
        if file.endswith(".yaml"):
            logging.info("Running job from {}".format(job))
            
            if RunJob(job, dry_run=args.dry_run):
                logging.debug("Job status OK {}".format(job))
                succeeded += 1
            else:
                logging.warning("Job status failed {}".format(job))
                failed += 1
        else:
            logging.warning("Ignoring job with incorrect extension: {}"
                    .format(job))

    logging.info("Finished: {} succeeded and {} failed"
            .format(succeeded, failed))

    # TODO: Option to eject/safely remove afterward?


if __name__ == "__main__":
    main()
