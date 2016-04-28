#!/usr/bin/env python

import officedissector
import argparse
import os

VERSION = "0.0.1"


def parseFile(filePath, outputDir):
    """"""

    print('Path: ' + filePath)
    doc = officedissector.doc.Document(filePath)
    print('Type: ' + doc.type)
    print('Is Template: ' + str(doc.is_template))
    print('Has Macro: ' + str(doc.is_macro_enabled))

    # DEBUG
    # print dir(officedissector)
    # print dir(doc)
    # print dir(doc.core_properties)

    print('\nFeatures:')
    if len(doc.features.comments) > 0:
        print('Comments:')
        print(doc.features.comments)
    if len(doc.features.custom_properties) > 0:
        print('Custom Properties:')
        print(doc.features.custom_properties)
    if len(doc.features.custom_xml) > 0:
        print('Custom XML:')
        print(doc.features.custom_xml)
    if len(doc.features.digital_signatures) > 0:
        print('Digital Signatures:')
        print(doc.features.digital_signatures)
    #if len(doc.features.doc) > 0:
    #    print doc.features.doc
    if len(doc.features.embedded_controls) > 0:
        print('Embedded Controls:')
        print(doc.features.embedded_controls)
    if len(doc.features.embedded_objects) > 0:
        print('Embedded Objects:')
        print(doc.features.embedded_objects)
    if len(doc.features.embedded_packages) > 0:
        print('Embedded Packages:')
        print(doc.features.embedded_packages)
    if len(doc.features.fonts) > 0:
        print('Fonts:')
        print(doc.features.fonts)
    if len(doc.features.images) > 0:
        print('Images:')
        print(doc.features.images)
    if len(doc.features.macros) > 0:
        print('Macros:')
        print(doc.features.macros)

        for m in doc.features.macros:
            #tempMacro = doc.part_by_name[m.name]
            outfile = open(os.path.join(outputDir, getSafeFileName(m.name)), "wb")
            outfile.write(m.stream().read())

    if len(doc.features.sounds) > 0:
        print('Sounds:')
        print(doc.features.sounds)
    if len(doc.features.videos) > 0:
        print('Videos:')
        print(doc.features.videos)

    print('\nCore Properties:')
    if len(doc.core_properties.category) > 0:
        print('Category: ' + doc.core_properties.category)
    if len(doc.core_properties.content_status) > 0:
        print('Content Status: ' + doc.core_properties.content_status)
    #if len(doc.core_properties.core_prop_part) > 0:
    #    print doc.core_properties.core_prop_part
    if len(doc.core_properties.created) > 0:
        print('Created: ' + doc.core_properties.created)
    if len(doc.core_properties.creator) > 0:
        print('Creator: ' + doc.core_properties.creator)
    if doc.core_properties.description is not None:
        print('Description: ' + doc.core_properties.description)
    if len(doc.core_properties.identifier) > 0:
        print('Identifier: ' + doc.core_properties.identifier)
    if len(doc.core_properties.keywords) > 0:
        print('Keywords: ' + doc.core_properties.keywords)
    if len(doc.core_properties.language) > 0:
        print('Language: ' + doc.core_properties.language)
    if len(doc.core_properties.last_modified_by) > 0:
        print('Last Modified By: ' + doc.core_properties.last_modified_by)
    if len(doc.core_properties.last_printed) > 0:
        print('Last Printed: ' + doc.core_properties.last_printed)
    if len(doc.core_properties.modified) > 0:
        print('Modified: ' + doc.core_properties.modified)
    if len(doc.core_properties.name) > 0:
        print('Name: ' + doc.core_properties.name)
    #if len(doc.core_properties.parse_all) > 0:
    #    print doc.core_properties.parse_all
    if len(doc.core_properties.revision) > 0:
        print('Revision: ' + doc.core_properties.revision)
    if doc.core_properties.subject is not None:
        print('Subject: ' + doc.core_properties.subject)
    if doc.core_properties.title is not None:
        print('Title: ' + doc.core_properties.title)
    if len(doc.core_properties.version) > 0:
        print('Version: ' + doc.core_properties.version)

    mp = doc.main_part()
    print('Main Part Content Type: ' + mp.content_type())
    print('Main Part Name: ' + mp.name)


def getSafeFileName(data):
    keepcharacters = (' ', '.', '_')
    return "".join(c for c in data if c.isalnum() or c in keepcharacters).rstrip()


def main():
    """Parse the command line parameters and load the configuration."""
    parser = argparse.ArgumentParser(description='Example: officefileinfo --file "/path/to/docx" ')
    parser.add_argument('-f', '--file', help='File path for MS Office file e.g. docx, xlsx', required=True)
    parser.add_argument('-o', '--outputdir', help='Output directory', required=True)
    args = parser.parse_args()

    print '\nofficefileinfo\n'

    print('Version: ' + VERSION)
    print('Company: Info-Assure')
    print('Author: Mark Woan\n')

    if os.path.exists(args.file) is False:
        print('The "file" parameter value does not exist')
        exit(1)

    if os.path.exists(args.outputdir) is True:
        if os.path.isdir(args.outputdir) is False:
            print('The "outputdir" parameter value is not a directory')
            exit(1)
    else:  # Directory does not exist so create it
        try:
            os.makedirs(args.outputdir)
        except OSError, e:
            print('Error creating output directory: ' + e.message)
            exit(1)

    parseFile(args.file, args.outputdir)

if __name__ == "__main__":
    main()