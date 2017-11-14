#Author-https://github.com/PopsDylan
#Description-Extract the BOM for a design to a CSV file

import adsk.core, adsk.fusion, traceback

# extract information from an occurence or a component
def extract_info(occ, comp):
    material = comp.material.name if comp.material else ('Various' if len(occ.childOccurrences) > 0 else 'Unknown')
    comp_appearance = comp.material.appearance.name if comp.material and comp.material.appearance else 'None'    
    
    info = {
        'Part Number': comp.partNumber,
        'Description': comp.description,
        'Material':    material,
        'Appearance':  occ.appearance.name if occ.appearance else comp_appearance
        }
    
    return info
    
# determine whether or not the occurrence is one that belongs in the BOM
def belongs_in_BOM(occ):
    # if the name starts with an underscore, exclude it
    #if occ.name.startswith('_'):
    # if any component of the full path starts with an underscore, exclude it
    if [name for name in occ.fullPathName.split('+') if name.startswith('_')]:
        return False
    # if it has no children, include it
    elif len(occ.childOccurrences) == 0:
        return True
    # if it has children and one of their names does not start with an underscore, exclude it
    else:
        for child in occ.childOccurrences:
            if not child.name.startswith('_'):
                return False
        return True
     
    return False
        
        
def extract_bom(design):
    root = design.rootComponent
    occurrences = root.allOccurrences
    
    # Gather information about each unique component
    bom = []
    bom_exclusions = set()
    
    for occ in occurrences:
        comp = occ.component
            
        if belongs_in_BOM(occ):
            info = extract_info(occ, comp)
            
            # entity is already in BOM?
            existing_bom_item = ([bom_item for bom_item in bom if bom_item['Component'] == comp and bom_item['Info'] == info]  or [None])[0]
        
            if existing_bom_item:
                existing_bom_item['Quantity'] += 1

            else:
                # Add this entity to the BOM
                bom.append({
                    'Component': comp,
                    'Name': comp.name,
                    'Quantity': 1,
                    'Info': info
                    })
        else:
            bom_exclusions.add(comp.name)
            
    return bom, bom_exclusions
    

    
def writeBOMToFile(bom, bom_exclusions):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        fileDlg = ui.createFileDialog()
        fileDlg.isMultiSelectEnabled = False
        fileDlg.title = 'BOM File Dialog'
        fileDlg.filter = '*.csv'

        dlgResult = fileDlg.showSave()
        if dlgResult == adsk.core.DialogResults.DialogOK:
            save_filename = fileDlg.filename   

            with open(save_filename,'w', newline='\r\n') as f:
                headings = ['Name', 'Quantity', 'Part Number', 'Description', 'Material','Appearance']
                item_fields = headings[:2]
                item_info_fields = headings[2:]
                
                delimiter = ','
                
                f.write(delimiter.join(headings) + '\n')
                
                for item in bom:
                    f.write(delimiter.join(['"{0}"'.format(str(item[field]).replace('"', '""')) for field in item_fields]))
                    f.write(delimiter)
                    f.write(delimiter.join(['"{0}"'.format(str(item['Info'][field]).replace('"', '""')) for field in item_info_fields]))
                    f.write('\n')
                    
                f.close()
                  
        
            with open(save_filename + '_exclusions.txt', 'w', newline='\r\n') as f1:
                for item in bom_exclusions:
                    f1.write(item + '\n')
                
                f1.close()
                
            msg = "Saved to \'" + save_filename + "\'"
            ui.messageBox(msg)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
    
def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'BOM Extractor'
        if not design:
            ui.messageBox('No active design', title)
            return
                    
        bom, bom_exclusions = extract_bom(design)
        
        writeBOMToFile(bom, bom_exclusions)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))



