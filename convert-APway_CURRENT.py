#! /usr/bin/python3


DEBUG = True
DEBUG = False


APWAY_SOURCE_FILE = './P2Way_edit.xml'
FAVORITES_ROOT = 'L:/2015_newnet/IE11/APplus/Favorites/APplus'
APPLUS_ENVIRONMENT = 'http://srvapprod/applusprod/'
APPLUS_ICON = '%windir%/_PE/APplus.ico'


LINK_TAGS = [ 'item', 'dlg', 'host', 'webfolder']


logfile = './test.log'
log = open( logfile, mode='tw', encoding='cp1252')


import os
import xml.etree.ElementTree as ElementTree


import win32com.client
WIN32COM_SHELL = win32com.client.Dispatch( 'WScript.Shell')


def main():
    tree = ElementTree.parse( APWAY_SOURCE_FILE)
    xml_things = tree.getroot()
    if not DEBUG:
        if not os.path.exists( FAVORITES_ROOT):
            os.mkdir( FAVORITES_ROOT)
    iter_dir = 0
    iter_file = 0
    for xml_thing in xml_things:
        iter_dir, iter_file = iterate_tree( xml_thing, FAVORITES_ROOT, iter_dir, iter_file)
    print( '\nDone!')


def iterate_tree( shit, current_path, iter_dir, iter_file):
    if clean_xsd(shit.tag) == 'folder':
        iter_dir += 1
        new_path = current_path + '/' + str(iter_dir*10).rjust( 3, '0') + ' - ' + shit.attrib['title'].strip().replace( '/', '-').replace( ': ', ' - ')
        make_directory( new_path)
        sub_iter_dir = 0
        sub_iter_file = 0
        for bull in shit:
            sub_iter_dir, sub_iter_file = iterate_tree( bull, new_path, sub_iter_dir, sub_iter_file)
    elif clean_xsd(shit.tag) in LINK_TAGS:
        iter_file += 1
        make_shortcut( current_path + '/', str(iter_file*10).rjust( 3, '0') + ' - ' + shit.attrib['title'].strip().replace( '/', '-').replace( ': ', ' - ') + '.lnk', sanitize_destination( shit.attrib['url']))
    else:
        print( 'ERROR:\nUNKNOWN tag = ' + clean_xsd(shit.tag) + '\n; attributes = ' + str(shit.attrib))
        exit()
    return( iter_dir, iter_file)


def sanitize_destination( destination):
    if destination[0:3] == '../':
        destination = APPLUS_ENVIRONMENT + destination[3:]
    elif destination[0:32] == "javascript:void(window.open('../":
        destination = APPLUS_ENVIRONMENT + destination[32:]
    elif destination[0:4] != "http":
        destination = APPLUS_ENVIRONMENT + destination
    else:
        pass
    return( destination)


def make_directory( path):
    log.write( 'make_directory(\t' + path + '\n')
    print( 'make_directory( "' + path + '")')
    if not DEBUG:
        if not os.path.exists( path):
            os.mkdir( path)


def make_shortcut( path, filename, destination):
    shortcut_pathfile = path + filename
    log.write( 'make_shortcut(\t' + shortcut_pathfile + '\t' + destination + '\n')
    print( 'make_shortcut(  "' + shortcut_pathfile + '", "' + destination + '")')
    if not DEBUG:
        try:
            createshortcut = WIN32COM_SHELL.CreateShortcut( shortcut_pathfile)
            createshortcut.TargetPath = destination
            createshortcut.IconLocation = APPLUS_ICON
            createshortcut.Save()
        except Exception as e:
            print( '==============================================================================')
            print( 'ERROR Exception:')
            print( '    WIN32COM_SHELL.CreateShortcut()')
            print( '        shortcut_pathfile = ' + shortcut_pathfile)
            print( '        TargetPath = ' + destination)
            print( '------------------------------------------------------------------------------')
            print( e)
            print( '==============================================================================')


def clean_xsd( string):
    return string.split( '}')[1]


main()

log.close()

