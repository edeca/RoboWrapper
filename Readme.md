# RoboWrapper manual #

RoboWrapper is a Python script that executes [Robocopy](https://en.wikipedia.org/wiki/Robocopy) with simple rule files.  These rule files allow auto-discovery of drive names and safety checks.

RoboWrapper is only designed to run on Windows.

If you were instead looking for robot rapping [see Youtube](https://www.youtube.com/watch?v=mvrva8NoMLM).

## Why? ##

I wrote primarily this because drive letters on Windows can change regularly unless fixed.  Sample use cases include:

 * Copying files to a USB stick that changes drive letter each time it is inserted. 
 * Syncing files to multiple USB sticks (where all have the same volume label).
 * Copying files from multiple computers to a network drive.

## Usage ##

For help, run: `python RoboWrapper.py --help`.

To run a job, run: `python RoboWrapper.py job.yaml` (add `-v` for debug).

To list drives, run: `python RoboWrapper.py --drives`.

## Basic configuration examples ##

The most basic configuration needs a name and source / destination paths:

```
name: Backup some stuff
source:
    path: 'C:\MyStuff'
destination:
    path: 'E:\Backup\MyStuff'
```

However this isn't very useful if the USB drive letter changes.  A more robust configuration would use drive serial numbers:

```
name: Backup some stuff
source:
    serial: '70DC81A0' # OS drive
    path: 'MyStuff'
destination:
    serial: '2C724B02' # Removable USB
    path: 'Backup\MyStuff'
```

Robocopy options can also be specified which will be passed when copying:

```
name: Backup some stuff
source:
    serial: '70DC81A0' # OS drive
    path: 'MyStuff'
destination:
    serial: '2C724B02' # Removable USB
    path: 'Backup\MyStuff'
robocopy:
	# Warning, /MIR can delete files!
    options: "/FFT /DST /MIR /XJ /NP /R:3 /W:30"
```

## Advanced configuration examples ##

### Copying to a specific volume name ###

Sometimes an exact serial may not be desirable.  For example the configuration below will copy to the first USB storage drive named "KINGSTON":

```
name: Sync to any drive named KINGSTON
source:
    serial: '70DC81A0' # OS drive
    path: 'MyStuff'
destination:
    # Note the first drive found is used.  If serial is also
    # specified then this will take precedence (if found).
    label: 'KINGSTON'
    path: 'Backup\MyStuff'
```

This is useful to duplicate the same files to a number of different storage devices which are plugged in one at a time, or if multiple users all have their own USB drive named "BACKUP".

### Checking a safety flag ###

To ensure files and folders are not overwritten by mistake an optional flag can be specified.  This should be a file (with any contents, or empty) which is checked before Robocopy is executed.

In the example below the files `C:\MyStuff\Robocopy.flag` and `E:\Robocopy.flag` must be present or Robocopy will not be run.

```
name: Backup my stuff, check a flag first
source:
    serial: '70DC81A0'     # OS drive
    path: 'MyStuff'        # Works out to C:\MyStuff\
    flag: '$path$\Robocopy.flag'
destination:
    serial: '2C724B02'     # Removable USB
    path: 'Backup\MyStuff' # Works out to E:\Backup\MyStuff\
    flag: '$drive$\Robocopy.flag'
```

The special variables `$drive$` and `$path$` will be substituted with the drive and path found by serial number or label.

### Using Windows environment variables ###

RoboWrapper will expand a number of Windows environment variables automatically.  The variables should be surrounded by dollar signs like the example below.

```
name: Backup AppData
source:
    path: '$APPDATA$'
destination:
    serial: '2C724B02' # Removable USB
    path: 'Backup\$COMPUTERNAME$\AppData'
```

Expansion in this way is preferable to the usual `%AppData%` format because Python can check the relevant directories exist before executing Robocopy.

### Adding a Robocopy log ###

The Robocopy log can be specified like below:

```
name: Backup some stuff
source:
    path: 'C:\MyStuff'
destination:
    path: 'E:\Backup\MyStuff'
robocopy:
    # Tell Robocopy to save a log
    log: '$dst_drive$\Robocopy-$timestamp$.log'
settings:
    # Optional format parameter for $timestamp$
    time_format: "%Y-%m-%dT%H%M%S"
```

The variables `$dst_drive$`, `$src_drive$`, `$dst_path$` and `$src_path$` will be expanded with the correct values.  `$timestamp$` will be substituted with the current time, which will be seconds since the epoch or may be formatted by strftime() using `time_format` (see [format strings](https://docs.python.org/2/library/datetime.html#strftime-and-strptime-behavior)).

## FAQ ##

### How can I test a job? ###

Add the `--dry-run` parameter when calling the script.  This will go through all the same steps of resolving drive letters, substituting environment variables and checking paths but will not execute Robocopy.

Run with `-v` for debug output.

### How can I get drive serial numbers? ###

Call the script with the `--drives` option which will display a table like below.

```
+-------+----------+------+
| Drive |  Serial  | Name |
+-------+----------+------+
|   C:  | 70DC81A0 |      |
|   D:  |   None   | None |
|   E:  | 2C724B02 | data |
+-------+----------+------+
```

## Requirements ##

The script requires [pywin32](http://sourceforge.net/projects/pywin32/) which needs to be installed using the setup wizard.

A number of other dependencies are also required, these can be installed using pip:

`pip install -r Requirements.txt` 