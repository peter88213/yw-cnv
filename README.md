# The yw-cnv extension for LibreOffice: Import and export yWriter 7 projects 

For more information, see the [project homepage](https://peter88213.github.io/yw-cnv) with description and download instructions.

## Development

*yw-cnv* depends on the [pywriter](https://github.com/peter88213/PyWriter) library which must be present in your file system. It is organized as an Eclipse PyDev project. The official release branch on GitHub is *main*.

### Mandatory directory structure for building the application script

```
.
├── PyWriter/
│   └── src/
│       └── pywriter/
└── yw-cnv/
    ├── src/
    ├── test/
    └── tools/ 
        ├── build.xml
        └── build_extension.py 
```

### Conventions

See https://github.com/peter88213/PyWriter/blob/main/docs/conventions.md

### Development tools

- [Python](https://python.org) version 3.10
- [Eclipse IDE](https://eclipse.org) with [PyDev](https://pydev.org) and *EGit*)
- Apache Ant for building the application script
- [pandoc](https://pandoc.org/) for building the HTML help pages

### Documentation tools

- [Eclipse Papyrus](https://www.eclipse.org/papyrus/) Modeling environment for creating Use Case and Class diagrams

## Credits

- [OpenOffice Extension Compiler](https://wiki.openoffice.org/wiki/Extensions_Packager#Extension_Compiler) by Bernard Marcelly.
- Frederik Lundh published the [xml pretty print algorithm](http://effbot.org/zone/element-lib.htm#prettyprint).
- Andrew D. Pitonyak published useful Macro code examples in [OpenOffice.org Macros Explained](https://www.pitonyak.org/OOME_3_0.pdf).

## License

This extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
