#Author-PopsDylan
#Description-Exports component bodies as STL.

import adsk.core, adsk.fusion, adsk.cam, traceback
import os.path



def collect_bodies_to_export(design):
    root = design.rootComponent
    occurrences = root.allOccurrences
    
    bodies_to_export = []    
    

    try:
        for comp in design.allComponents:
            supplier, partnumber = [s.strip() for s in (comp.partNumber.split(':', 1) if ':' in comp.partNumber else ('Unknown', comp.partNumber))]

            if supplier == 'Printed Part':
                occurrence_count = root.allOccurrencesByComponent(comp).count

                if occurrence_count != 0:
                    bodies = comp.bRepBodies
            
                    if bodies.count == 1:
                        # add to collection for export
                        filename = str(partnumber)
                        
                        if occurrence_count > 1:
                            filename += "_x" + str(occurrence_count)
                        bodies_to_export.append ((bodies.item(0), filename))
                    else:
                        raise Exception('Component ' + comp.name + ' is a printed part with ' + str(bodies.count) + ' bodies.')
    
    except:
        app = adsk.core.Application.get()
        ui  = app.userInterface   
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            
        bodies_to_export = []
        
    return bodies_to_export
    
def export_stls(design):
    bodies_to_export = collect_bodies_to_export(design)
        
    app = adsk.core.Application.get()
    ui  = app.userInterface    

    if bodies_to_export:
        folderDlg = ui.createFolderDialog();
        folderDlg.title = 'STL Export Folder'

        dlgResult = folderDlg.showDialog()
    
        if dlgResult == adsk.core.DialogResults.DialogOK:
            save_folder = folderDlg.folder
            
            exportManager = design.exportManager
        
            for body, filename in bodies_to_export:
                filepath = os.path.join(save_folder, filename) 
                stlExportOptions = exportManager.createSTLExportOptions(body, filepath)
                stlExportOptions.sendToPrintUtility = False
                stlExportOptions.isBinaryFormat = False
                
                exportManager.execute(stlExportOptions)
                
            if ui:
                ui.messageBox('Exported ' + str(len(bodies_to_export)) + ' STL files to ' + save_folder)
    else:
        if ui:
            ui.messageBox('Found no printed parts to export.')
    
def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
     
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'STLExporter'
        if not design:
            ui.messageBox('No active design', title)
            return
            
        export_stls(design)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
