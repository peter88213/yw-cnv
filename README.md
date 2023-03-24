# The yw-cnv extension for LibreOffice: Import and export yWriter 7 projects 

For more information, see the [project homepage](https://peter88213.github.io/yw-cnv) with description and download instructions.

## Development

*yw-cnv* is organized as an Eclipse PyDev project. The official release branch on GitHub is *main*.

### Conventions

See https://github.com/peter88213/PyWriter/blob/main/docs/conventions.md

Exceptions:
- The directory structure is modified to minimize dependencies:

```
.
└── yw-cnv/
    ├── src/
    ├── test/
    └── tools/ 
        └── build.xml
```

### Development tools

- [Python](https://python.org) version 3.9
- [Eclipse IDE](https://eclipse.org) with [PyDev](https://pydev.org) and [EGit](https://www.eclipse.org/egit/)
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
