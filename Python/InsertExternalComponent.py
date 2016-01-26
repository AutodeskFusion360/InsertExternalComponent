#Author-
#Description-

import adsk.core, adsk.fusion, traceback

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

commandId = 'InsertExternalComponentCmd'
commandDescription = 'Insert External Component'
projectInputId = commandId + '_project'
fileInputId = commandId + '_file'
app = None
ui = None
projectFiles = {}
projects = None

def getFile(projectName, fileName):
    project = getProject(projectName)    
    
    if project:
        dataFiles = project.rootFolder.dataFiles
        for file in dataFiles:
            if file.name == fileName:
                return file
            
    return None

def getProject(projectName):
    for project in projects:
        if project.name == projectName:
            return project
            
    return None

# returns first project's name
def fillProjectsDictionary():
    firstProject = None
    for project in projects:
        if not firstProject:
            firstProject = project
            
        global projectFiles
        projectFiles[project.name] = None 
        
    return firstProject.name
        
def fillFilesDictionary(projectName):
    files = projectFiles[projectName]
    if not files:
        project = getProject(projectName)
        dataFiles = project.rootFolder.dataFiles
        files = {}
        for file in dataFiles:
            files[file.name] = file  
        
        projectFiles[projectName] = files

def addItemsToDropdown(items, dropdownInput):
    dropdownItems = dropdownInput.listItems
    
    # gather items to delete
    itemsToDelete = []
    for dropdownItem in dropdownItems:
        itemsToDelete.append(dropdownItem)
    
    # add new items 
    firstNewItem = None
    for item in items:
        newItem = dropdownItems.add(item, False, '')
        if not firstNewItem:
            firstNewItem = newItem
            firstNewItem.isSelected = True
    
    # delete existing items
    for dropdownItem in itemsToDelete:
        dropdownItem.deleteMe()   
        
    return firstNewItem

# Event handler for inputChanged event.
class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        # Code to react to the event.
        commandInput = args.input
        if commandInput.id == projectInputId:
            projectName = commandInput.selectedItem.name
            fillFilesDictionary(projectName)
            
            currentProject = projectFiles[projectName]
            
            fileInput = commandInput.commandInputs.itemById(fileInputId)            
            addItemsToDropdown(currentProject, fileInput)     

class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:  
            command = args.firingEvent.sender 
            projectName = command.commandInputs.itemById(projectInputId).selectedItem.name
            fileName = command.commandInputs.itemById(fileInputId).selectedItem.name    
            file = getFile(projectName, fileName) 
                
            design = app.activeProduct  
            root = design.rootComponent
            occs = root.occurrences
            occs.addByInsert(file, adsk.core.Matrix3D.create(), True)   

        except:
            if ui:
                ui.messageBox('Command execution failed:\n{}'.format(traceback.format_exc()))

class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__() 
    def notify(self, args):
        try:
            # document needs to be saved before you can insert 
            # an external component in it
            doc = app.activeDocument
            if not doc.isSaved:
                 ui.messageBox("Document needs to be saved before you can insert a component in it.")
                 adsk.terminate()
                 return 
            
            global projects
            projects = app.data.dataProjects
                                         
            cmd = args.command
            cmd.setDialogInitialSize(300, 200)
            
            # handle command execute
            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute) 
            handlers.append(onExecute)
            
            # handle destroy
            onDestroy = CommandDestroyHandler()
            cmd.destroy.add(onDestroy)
            handlers.append(onDestroy)            
            
            # handle input changed
            onInputChanged = InputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)
        
            # add projects
            fillProjectsDictionary()
            dropdownInput = cmd.commandInputs.addDropDownCommandInput(projectInputId, 'Project', adsk.core.DropDownStyles.TextListDropDownStyle);
            projectItem = addItemsToDropdown(projectFiles, dropdownInput)
            
            # add files
            fillFilesDictionary(projectItem.name)
            dropdownInput = cmd.commandInputs.addDropDownCommandInput(fileInputId, 'File', adsk.core.DropDownStyles.TextListDropDownStyle);            
            project = projectFiles[projectItem.name]            
            addItemsToDropdown(project, dropdownInput)
            
        except:
            if ui:
                ui.messageBox('Command creation failed:\n{}'.format(traceback.format_exc()))

class CommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # when the command is done, terminate the script
            # this will release all globals which will remove all event handlers
            
            adsk.terminate()
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
                
def run(context):
    ui = None
    try:
        global app
        app = adsk.core.Application.get()
        global ui
        ui  = app.userInterface
        
        commandDefinitions = ui.commandDefinitions 
        
        commandDefinition = commandDefinitions.itemById(commandId)
        if not commandDefinition:
            commandDefinition = commandDefinitions.addButtonDefinition(commandId, commandDescription, commandDescription, '')
        onCommandCreated = CommandCreatedEventHandlerPanel()
        commandDefinition.commandCreated.add(onCommandCreated)
        # keep the handler referenced beyond this function
        handlers.append(onCommandCreated)
        
        commandDefinition.execute()

        # prevent this module from being terminate when the script returns, 
        # because we are waiting for event handlers to fire
        adsk.autoTerminate(False)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
