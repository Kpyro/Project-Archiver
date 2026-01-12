# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#  Copyright (c) 2020 by Patrick Rainsberry.
#  :license: Apache2, see LICENSE for more details.
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#  ExportCommand.py
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

import os

import adsk.core
import adsk.fusion
import adsk.cam

import apper
import config

SKIPPED_FILES = []


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def get_app_objects():
    app = adsk.core.Application.get()
    ui = app.userInterface
    design = adsk.fusion.Design.cast(app.activeProduct)
    return app, ui, design


def export_folder(root_folder, output_folder, file_types, write_version, name_option, folder_preserve):

    app, ui, design = get_app_objects()

    for folder in root_folder.dataFolders:

        if folder_preserve:
            new_folder = os.path.join(output_folder, folder.name, "")
            if not os.path.exists(new_folder):
                os.makedirs(new_folder)
        else:
            new_folder = output_folder

        export_folder(folder, new_folder, file_types, write_version, name_option, folder_preserve)

    for file in root_folder.dataFiles:
        if file.fileExtension == "f3d":
            open_doc(file)
            try:
                output_name = get_name(write_version, name_option)
                export_active_doc(output_folder, file_types, output_name)

            except Exception as e:
                ui.messageBox(str(e))
                break


def open_doc(data_file):
    app = adsk.core.Application.get()
    try:
        document = app.documents.open(data_file, True)
        if document:
            document.activate()
    except:
        pass


def export_active_doc(folder, file_types, output_name):
    global SKIPPED_FILES

    app, ui, design = get_app_objects()
    export_mgr = design.exportManager

    export_functions = [
        export_mgr.createIGESExportOptions,
        export_mgr.createSTEPExportOptions,
        export_mgr.createSATExportOptions,
        export_mgr.createSMTExportOptions,
        export_mgr.createFusionArchiveExportOptions,
        export_mgr.createSTLExportOptions
    ]

    export_extensions = ['.igs', '.step', '.sat', '.smt', '.f3d', '.stl']

    for i in range(file_types.count - 2):
        if file_types.item(i).isSelected:
            export_name = dup_check(folder + output_name + export_extensions[i])
            export_options = export_functions[i](export_name)
            export_mgr.execute(export_options)

    # F3D (no external references)
    if file_types.item(file_types.count - 2).isSelected:
        if app.activeDocument.allDocumentReferences.count > 0:
            SKIPPED_FILES.append(app.activeDocument.name)
        else:
            export_name = dup_check(folder + output_name + '.f3d')
            export_options = export_mgr.createFusionArchiveExportOptions(export_name)
            export_mgr.execute(export_options)

    # STL
    if file_types.item(file_types.count - 1).isSelected:
        stl_name = folder + output_name + '.stl'
        stl_options = export_mgr.createSTLExportOptions(design.rootComponent, stl_name)
        export_mgr.execute(stl_options)


def dup_check(name):
    if os.path.exists(name):
        base, ext = os.path.splitext(name)
        return dup_check(base + '-dup' + ext)
    return name


def get_name(write_version, option):
    app, ui, design = get_app_objects()
    output_name = ''

    if option == 'Document Name':
        doc_name = app.activeDocument.name
        if not write_version and ' v' in doc_name:
            doc_name = doc_name[:doc_name.rfind(' v')]
        output_name = doc_name

    elif option == 'Description':
        output_name = design.rootComponent.description

    elif option == 'Part Number':
        output_name = design.rootComponent.partNumber

    else:
        raise ValueError('Invalid name option')

    return output_name


def update_name_inputs(command_inputs, selection):
    command_inputs.itemById('write_version').isVisible = (selection == 'Document Name')


# -----------------------------------------------------------------------------
# Command Class
# -----------------------------------------------------------------------------

class ExportCommand(apper.Fusion360CommandBase):

    def on_input_changed(self, command, inputs, changed_input, input_values):
        if changed_input.id == 'name_option_id':
            update_name_inputs(inputs, changed_input.selectedItem.name)

    def on_execute(self, command, inputs, args, input_values):
        global SKIPPED_FILES

        app, ui, design = get_app_objects()

        output_folder = input_values['output_folder']
        folder_preserve = input_values['folder_preserve_id']
        file_types = inputs.itemById('file_types_input').listItems
        write_version = input_values['write_version']
        name_option = input_values['name_option_id']

        root_folder = app.data.activeProject.rootFolder

        if not output_folder.endswith(os.path.sep):
            output_folder += os.path.sep

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        export_folder(
            root_folder,
            output_folder,
            file_types,
            write_version,
            name_option,
            folder_preserve
        )

        if SKIPPED_FILES:
            ui.messageBox(
                "The following files contained external references and could not be exported as f3d's:\n\n{}".format(
                    '\n'.join(SKIPPED_FILES)
                )
            )

        close_command = ui.commandDefinitions.itemById(
            self.fusion_app.command_id_from_name(config.close_cmd_id)
        )
        close_command.execute()

    def on_create(self, command, inputs):
        global SKIPPED_FILES
        SKIPPED_FILES.clear()

        default_dir = apper.get_default_dir(config.app_name)
        inputs.addStringValueInput('output_folder', 'Output Folder:', default_dir)

        drop_input = inputs.addDropDownCommandInput(
            'file_types_input',
            'Export Types',
            adsk.core.DropDownStyles.CheckBoxDropDownStyle
        )

        items = drop_input.listItems
        items.add('IGES', False)
        items.add('STEP', True)
        items.add('SAT', False)
        items.add('SMT', False)
        items.add('F3D', False)
        items.add('STL', False)

        name_option = inputs.addDropDownCommandInput(
            'name_option_id',
            'File Name Option',
            adsk.core.DropDownStyles.TextListDropDownStyle
        )

        name_option.listItems.add('Document Name', True)
        name_option.listItems.add('Description', False)
        name_option.listItems.add('Part Number', False)

        inputs.addBoolValueInput(
            'folder_preserve_id',
            'Preserve folder structure?',
            True,
            '',
            True
        )

        inputs.addBoolValueInput(
            'write_version',
            'Write versions to output file names?',
            True,
            '',
            False
        ).isVisible = False

        update_name_inputs(inputs, 'Document Name')