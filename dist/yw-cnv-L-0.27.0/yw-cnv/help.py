import os

def show_help():
    scriptLocation = os.path.dirname(__file__)
    os.startfile(scriptLocation + '/help.html')
    
def show_adv_help():
    scriptLocation = os.path.dirname(__file__)
    os.startfile(scriptLocation + '/help_adv.html')