# visual basic 6 plugin
Jenkins plugin for Visual Basic 6 builder

Automating a build of a Visual Basic 6 has some challenges. This plugin aims to workaround some of these issues. 

# Usage
In Jenkins Configure System, section VB6 Builder, set the VB6.exe path. 

![ScreenShot](VB6_path.png?raw=true )

In a job configuration add a VB6 build step and define the path to the project file.  

![ScreenShot](job_config.png?raw=true)

## See also
http://zbz5.net/automating-build-visual-basic-6-project

## Common File Extensions in a Visual Basic 6 Project
```
Extension   Description
.frm    Form
.bas    Module
.frx    Automatically generated file for every graphic in the project
.ocx    ActiveX Control
.cls    Class Module
.vbp    Visual Basic Project
.ctlUser    Control File
.ctxUser    Control Binary File
.dca    Active Designer Cache
.ddf    Package and Deployment Wizard CAB Information File
.dep    Package and Deployment Wizard Dependency File
.dob    ActiveX Document Form File
.dox    ActiveX Document Binary Form File
.dsr    Active Designer File
.dsx    Active Designer Binary File
.dws    Deployment Wizard Script File
.log    Log file for Load Errors
.oca    Control TypeLib Cache File
.pag    Property Page File
.pgx    Binary Property Page File
.res    Resource File
.tlb    Remote Automation TypeLib File
.vbg    Visual Basic Group Project File
.vbl    Control Licensing File
.vbr    Remote Automation Registration File
.vbw    Visual Basic Project Workspace File
.vbz    Wizard Launch File
.wct    WebClass HTML Template
```

## Flowchart

![Mermaid](https://raw.githubusercontent.com/brunocantisano/visual-basic-6-plugin/master/mermaid-diagram-vb6-plugin.svg?sanitize=true)
<img src="https://raw.githubusercontent.com/brunocantisano/visual-basic-6-plugin/master/mermaid-diagram-vb6-plugin.svg?sanitize=true">
