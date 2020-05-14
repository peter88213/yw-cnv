# PyWriter tools for LibreOffice: Edit yWriter scenes and metadata with LibreOffice

Export an yWriter 6/7 project to an OpenDocument file with invisible chapter and scene markers or to an csv file. 
Write back the edited data to the yWriter project file.

## Features

### yWriter5 style proof reading

Export an yWriter 6/7 project to an OpenDocument file with chapter and scene markers. 
Write back the proofread scenes to the yWriter 6/7 project file.

### Project export to OpenDocument format

Generate a "standard manuscript" formatted OpenDocument textfile from an yWriter 6/7 project.

### Convert the manuscript without visible chapter/scene tags

Export an yWriter 6/7 project to an OpenDocument file with invisible chapter and scene sections (to be seen in the Navigator). 
Write back the proofread scenes to the yWriter 6/7 project file.

### Generate editable reports

* Very brief synopsis (Part descriptions): Chapter titles and chapter descriptions that can be edited in Office Writer and written back to yWriter format. Applies only to chapters marked _This chapter begins a new section_ in yWriter. 

* Brief synopsis (Chapter descriptions): Chapter titles and chapter descriptions that can be edited in Office Writer and written back to yWriter format. Doesn't apply to chapters marked _This chapter begins a new section_ in yWriter.

* Full synopsis (Scene descriptions): Chapter titles and scene descriptions that can be edited in Office Writer and written back to yWriter format.

* Character sheets that can be edited in Office Writer and written back to yWriter format (yet to come). 

* World building sheets containing locations and items that can be edited in Office Writer and written back to yWriter format (yet to come). 

### Provide a scene list

Create (via csv) a spreadsheet table listing scene title, scene descriptions, and links to the manuscript's scene sections. This list can be enhanced with scene metadata (e.g. tags, goals, time).

### Provide a plot list

Create (via csv) spreadsheet tables containing plot related metadata that can be displayed and edited.

#### Plotting conventions

In yWriter, you can divide your novel into  __Plot Sections__  (e.g. acts or steps) by inserting chapters marked "other". They will show up in blue color and won't get exported.


__Plot-related events__ (e.g. "Mid Point", "Climax") can be identified by "scene tags" if you want to link them to a specific scene.


You can use scene notes for  __plot-specific explanations__. 


If you want to  __visualize character arcs__, you can use the project's rating names by changing them to the names of up to four main characters. Then you can quantify the state of these four characters and put them into the scenes. It's easy then to let LibreOffice Calc show a diagram for the scene ratings over scene count or word count.

* Only rating field names corresponding to character names or containing the string "line", e.g. "A-Line", "BStoryline" (up to 10 case insensitive characters) appear in the plot list.
* Only ratings greater than 1 appear in the plot list, i.e. 1 means "a rating is not set for this chapter". 
* Recommended ratings: 
    * 1 = N/A
    * 2 = unhappy
    * 3 = dissatisfied
    * 4 = vague
    * 5 = satisfied
    * 6 = happy

* Ratings deleted while editing the plotlist will be converted to 1 on writing back. 

## Requirements

* Windows.

* yWriter 6 or yWriter 7.

* A LibreOffice 5 or 6 standard installation (not a "portable" version).

## Download

The  _PyWriter Tools_  Software comes as a zipfile `PyWriter_LO_<version number>.zip`. 

[Download page](https://github.com/peter88213/pywlo/releases/latest)



## How to install

1. Unzip `PyWriter_LO_<version number>.zip` within your user profile.

2. Move into the `PyWriter_LO_<version number>` folder and run `Install.bat` (double click).
   This will copy all needed files to the right places, install a LibreOffice extension.
   You may be asked for approval to modify the Windows registry. Please accept in order to 
   install an Explorer context menu entry "PyWriter Tools" for yWriter 6/7 files.

3. Start LibreOffice Writer. You should see a small toolbar window containing a button with
   a yWriter logo. This button is for writing the proofread file back to the yWriter project.
   Dock this toolbar anywhere for your convenience. Switch to LibreOffice Calc and repeat the steps described above.

5. __Optional:__  Download and install the [Courier Prime](https://quoteunquoteapps.com/courierprime).



## How to use

1. Write your novel with yWriter. Please consider the following conventions:
    * Text markup: Bold and italics are supported. Other highlighting such as underline and strikethrough are lost.
    * All chapters and scenes will be exported, whether "used" or "unused". 
    * If `This chapter begins a new section` is selected in _Chapter/Details_, the heading will be on the first level. Otherwise, it will be on the second level.
    * If `Suppress chapter title when exporting` is selected in _Chapter/Details_, PyWriter Tools will remove "Chapter" from auto-numbered chapter titles. The numbers will remain. These modifications have no effect on the reimport.

   Backup entire project and close yWriter.

2.  Move into your yWriter project folder, and right-click your .yw6 or .yw7 project file. 
   In the context menu, choose `Proof read with LibreOffice`. 
   
![Screenshot: Windows Explorer context menu](https://raw.githubusercontent.com/peter88213/PyWriter Tools/master/docs/Screenshots/PyWriter Tools_cm.png)

3. If everything goes well, you find an OpenDocument file named `<your yWriter project>_<type>.odt`, or `<your yWriter project>_<type>.csv`. Open it (double click) for editing. 
    * If your standard application for `.csv` Files is not LibreOffice, please right-click the .csv file and choose `open with LibreOffice Calc`. You will be prompted for csv conversion settings. Please select `utf-8` encoding and `|` as field separator.
    * The ODT text documents contain Chapter `[ChID:x]`
   and/or scene `[ScID:y]` sections. You can see them in the Navigator window (F5).  __Do not touch the section structure__  if you want to reimport the document into yWriter. 

4. In order to write back the edited text to the yWriter project, klick the toolbar button
   with the yWriter logo, or select the menu item 
   `Tools > Add-ons > PyWriter Tools > Write back to yWriter`.

If everything went well, a window will open with a success message.



## How to uninstall

Move into the `PyWriter_LO_<version number>` folder and run `Uninstall.bat` (double click). 
You may be asked for approval to modify the Windows registry. Please accept in order to 
remove the Explorer context menu entry. 


_PyWriter tools_  are distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).