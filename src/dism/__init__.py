# DISM Module
# Short description : Easily use Microsoft's DISM Command-Line Interface in Python.
# Version : 0.0.2
# Made By : SamuelLouf <https://github.com/samuellouf>
# Creation date : 11th July 2024
# GitHub Page <https://github.com/samuellouf/dism>
# ------------------------------------------------------------------------------
# License : MIT
# Copyright 2024 SamuelLouf
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
# documentation files (the “Software”), to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
# and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial
# portions of the Software.
#
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
# TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
# OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
import subprocess, ctypes, sys, datetime

__version__ = '0.0.2'

# Python functions
def isUserAdmin():
  """ Returns True if the program has admin rights. """
  return not ctypes.windll.shell32.IsUserAnAdmin() == 0

def giveAdminRightsToProgram():
  """ Tries to get admin rights then return True if the attempt was a success. """
  if not isUserAdmin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    return isUserAdmin()

def init():
  """ Initializes the module """
  giveAdminRightsToProgram()

# DISM functions
# Basic usage of DISM
def checkHealth():
  """
    Detect component store corruption. 
  
    Returns :
        bool: Was a component store corruption detected ?
  """
  if not isUserAdmin(): raise OSError('No admin access')
  r = subprocess.run(['dism', '/Online', '/Cleanup-Image', '/CheckHealth', '/English'], capture_output=True, text=True)
  if 'No component store corruption detected.' in r.stdout:
    return True
  else :
    return False

def scanHealth():
  """
    Scan component store corruption. (/!\This operation can take time) 
  
    Returns :
        tuple: Was a component store corruption detected?, Was the opertation a success?
  """
  if not isUserAdmin(): raise OSError('No admin access')
  r = subprocess.run(['dism', '/Online', '/Cleanup-Image', '/ScanHealth', '/English'], capture_output=True, text=True)

  if 'No component store corruption detected.' in r.stdout:
    csc = True
  else :
    csc = False

  if 'The operation completed successfully.' in r.stdout:
    success = True
  else:
    success = False

  return (csc, success)

def restoreHealth():
  """
    Scans for component store corruption, performs repair operations automatically. (/!\This operation can take time) 
  
    Returns :
        tuple: Was a component store corruption detected?, Was the opertation a success?
  """
  if not isUserAdmin(): raise OSError('No admin access')
  r = subprocess.run(['dism', '/Online', '/Cleanup-Image', '/RestoreHealth', '/English'], capture_output=True, text=True)

  return ('The restore operation completed successfully.' in r.stdout)

def restoreHealthExt(external_source, limit_access = False):
  """
    Scans for component store corruption, performs repair operations automatically. The corruption cannot be repaired, use an external source. (/!\This operation can take time) 

    Args :
        external_source (str): Extrenal source to repair corruptions. 
        limit_access (bool): To stop DISM from using Windows Update and force it to use the external source for repairs.

    Returns :
        tuple: Was a component store corruption detected?, Was the opertation a success?
  """
  if not isUserAdmin(): raise OSError('No admin access')
  cmd = ['dism', '/Online', '/Cleanup-Image', '/RestoreHealth', '/Source:' + external_source, '/English']
  if limit_access:
    cmd.append('/LimitAccess')
  r = subprocess.run(cmd, capture_output=True, text=True)

  if 'No component store corruption detected.' in r.stdout:
    return True
  else :
    return False

# Working with Windows Images
def getWimInfo(wim, index = 1):
  """
    Get infos about wim file.

    Args:
        wim (str): Path to wim file

    Returns:
        data: Index, Name, Description and Size of the wim file
    """
  if not isUserAdmin(): raise OSError('No admin access')
  r = subprocess.run(['dism', '/English', '/Get-WimInfo', '/wimfile=' + wim, '/Index:' + str(index)], capture_output=True, text=True)
  
  class data:
    def __init__(self):
      self.index = self.size = self.version = self.ServicePack_build = self.ServicePack_level = self.directories = self.files = -1
      self.name = self.description = self.architecture = self.hal = self.edition = self.installation = self.productType = self.productSuite = self.systemRoot = self.defaultLanguage = None
      self.bootable = False
      self.created = self.modified = datetime.datetime.now()
      self.languages = []

  def turnStringIntoDate(string):
    datetime_ = string.split(' - ')[0].split('/') + string.split(' - ')[1].split(':')
    for x in range(datetime_.__len__()):
      datetime_[x] = int(datetime_[x])
    return datetime.datetime(*(datetime_[2], datetime_[1], datetime_[0], datetime_[3], datetime_[4], datetime_[5]))
  
  def getLanguages(output, default = False):
    languages = []
    lines = output.splitlines()

    for i in range(lines.__len__()):
      if 'Languages :' in lines[i]:
        for j in range(lines.__len__() - i):
          if ('The operation completed successfully.' not in lines[j + i]) and ('Languages :' not in lines[j + i]):
            lang = lines[j + i].replace(' ', '').replace('\t', '')

            if default == False:
              if '(Default)' in lang:
                default = lang.replace('(Default)', '')
              lang = lang.replace('(Default)', '')

            if lang != '':
              languages.append(lang)

    if default == False:          
      return languages
    else:
      return (languages, default)
  
  def getDefaultLanguage(output):
    return getLanguages(output, True)[1]

  data_ = data()

  for line in r.stdout.splitlines():
    if ' : ' in line:
      one = line.split(' : ')[1]
      match line.split(' : ')[0]:
        case 'Index':
          data_.index = int(one)
        case 'Name':
          data_.name = one
        case 'Description':
          data_.description = one
        case 'Size':
          data_.size = int(one.replace('ÿ', '').replace(' bytes', ''))
        case 'WIM Bootable':
          data_.bootable = (one == 'Yes')
        case 'Architecture':
          data_.architecture = one
        case 'Hal':
          data_.hal = one
        case 'Version':
          data_.version = one
        case 'ServicePack Build':
          data_.ServicePack_build = int(one)
        case 'ServicePack Level':
          data_.ServicePack_level = int(one)
        case 'Edition':
          data_.edition = one
        case 'Installation':
          data_.installation = one
        case 'ProductType':
          data_.productType = one
        case 'ProductSuite':
          data_.productSuite = one
        case 'System Root':
          data_.systemRoot = one
        case 'Directories':
          data_.directories = int(one)
        case 'Files':
          data_.files = int(one)
        case 'Created':
          data_.created = turnStringIntoDate(one)
        case 'Modified':
          data_.modified = turnStringIntoDate(one)
        case _:
          pass

  data_.languages = getLanguages(r.stdout)
  data_.defaultLanguage = getDefaultLanguage(r.stdout)

  return data_

def mountWim(wim, mounting_path, readonly=True):
  """
    Easily mount a wim file

    Args:
        wim (str): Path to wim file
        mounting_path (str): Path in which the wim file will be mounted
        readonly (bool): Mount the wim in readonly mode
    """
  if not isUserAdmin(): raise OSError('No admin access')
  cmd = ['dism', '/Mount-Wim', '/WimFile:' + wim, '/Index:1', '/MountDir:' + mounting_path]
  if readonly:
    cmd.append('/ReadOnly')
  r = subprocess.run(cmd, capture_output=True, text=True)
  return r

def unmountWim(mounting_path, commit=False):
  """
    Easily unmount a wim file

    Args:
        mounting_path (str): Path in which the wim file will be mounted
        commit (bool): Commit changes
    """
  if not isUserAdmin(): raise OSError('No admin access')
  cmd = ['dism', '/Unmount-Wim', '/MountDir:' + mounting_path]
  if commit:
    cmd.append('/Commit')
  else:
    cmd.append('/Discard')
  r = subprocess.run(cmd, capture_output=True, text=True)
  return r

def addPackage(mounting_path, package):
  """
    Easily add a package to a wim file

    Args:
        mounting_path (str): Path in which the wim file was mounted
        package (str): Path to package to add to the mounted wim
  """
  if not isUserAdmin(): raise OSError('No admin access')
  subprocess.run(['dism', '/Image:' + mounting_path, '/Add-Package', '/PackagePath:' + package])

def removePackage(mounting_path, package):
  """
    Easily remove a package to a wim file

    Args:
        mounting_path (str): Path in which the wim file was mounted
        package (str): Name of the package to remove from the mounted wim
  """
  if not isUserAdmin(): raise OSError('No admin access')
  subprocess.run(['dism', '/Image:' + mounting_path, '/Remove-Package', '/PackageName:' + package])

def addDriver(mounting_path, driver):
  """
    Easily add a driver to a wim file

    Args:
        mounting_path (str): Path in which the wim file was mounted
        driver (str): Path to driver to add to the mounted wim
  """
  if not isUserAdmin(): raise OSError('No admin access')
  subprocess.run(['dism', '/Image:' + mounting_path, '/Add-Driver', '/Driver:' + driver])

def removeDriver(mounting_path, driver):
  """
    Easily remove a driver from a wim file

    Args:
        mounting_path (str): Path in which the wim file was mounted
        driver (str): Path to driver to remove from the mounted wim
  """
  if not isUserAdmin(): raise OSError('No admin access')
  subprocess.run(['dism', '/Image:' + mounting_path, '/Add-Driver', '/Driver:' + driver])

# Capture and Apply Windows Images

def captureImage(wim, capture_dir, name):
  """ Capture an image of a Windows partition

      Args :
          wim (str): Output wim file
          capture_dir (str): To-be-captured partition's path
          name (str): Image's name

      Return :
        bool: Was the operation a success?
  """
  r = subprocess.run(['dism', '/Capture-Image', '/ImageFile:' + wim, '/CaptureDir:' + capture_dir, '/Name:"' + name + '"', '/English'], capture_output=True, text=True)
  return ('The operation completed successfully.' in r.stdout)

def applyImage(wim, apply_dir):
  """ Capture an image of a Windows partition

      Args :
          wim (str): Output wim file
          capture_dir (str): To-be-captured partition's path
          name (str): Image's name

      Return :
        bool: Was the operation a success?
  """
  r = subprocess.run(['dism', '/Apply-Image', '/ImageFile:' + wim, '/Index:1', '/ApplyDir:' + apply_dir, '/English'], capture_output=True, text=True)
  return ('The operation completed successfully.' in r.stdout)

# END