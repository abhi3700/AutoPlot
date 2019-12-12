# AutoPlot

<p align="center">
  <img src="../images/autoplot_wallpaper.jpg" alt="AutoPlot Wallpaper" width="" height="">
</p>

<div style="page-break-after: always;"></div>

## Table of Contents
* Installation
* Implementation
* Chart Features
* Troubleshooting

<div style="page-break-after: always;"></div>

## Installation
Both the Softwares doesn't need any `ADMIN` permission.

### Anaconda
* Download and Install [Anaconda](https://www.anaconda.com/distribution/#download-section). Also, kept in the local directory [here]().

> NOTE: Don't forget to tick the command terminal option (which is NOT recommended) in the last prompt.

<p align="center">
  <img src="../videos/anaconda_installation.gif" alt="Anaconda installation" width="" height="">
</p>

* Install packages
	```markdown
	pywin32
	xlwings
	retrying
	plotly
	```

<p align="center">
  <img src="../videos/anaconda_installation.gif" alt="Package Installation" width="" height="">
</p>

* Enable macro settings in MS Excel's options menu.
* Integrate `xlwings` with Excel. It would appear in separate tab inside Excel application.

Now, Excel is Integrated with Python.

### Chart auto-updation
Due to the incompatability issue of Windows Operating System in local Intranet, the charts (which opens in the browser) doesn't create/update the charts in the same directory.

So, [Fork](https://git-fork.com/) application needs to be installed for executing the shell (`.sh`) [Analogous to `.bat` in Windows] file. Also, kept in the local directory [here](https://www.anaconda.com/distribution/#download-section).



<div style="page-break-after: always;"></div>

## Implementation
### QC charts
This has been implemented in different case studies like __Dry Etch__, __Wet Etch__, __Diffusion__ so far.

1. ##### M-1: `button-mode` - Open the Excel file and press <kbd>RUN</kbd> button present in __RUN_code__ sheet. This opens the charts in the default Browser.

	> NOTE: Unfortunately, due to Windows incompatibility for local intranet during execution, the charts are not updated in the same directory. For this, follow __M-2__ 

<!-- TODO: Add .gif for this -->

1. ##### M-2: `shell-mode`
<!-- TODO: Add .gif for this -->

<!-- ### Control limits calculation -->
<!-- ### Wafer Map -->

<div style="page-break-after: always;"></div>

## Chart Features
* Show closest data on hover
<!-- TODO: Add .gif for this -->
* Compare data on hover
<!-- TODO: Add .gif for this -->
* See __"Remarks"__ in the chart
<!-- TODO: Add .gif for this -->
* Zoom-in/out from the center of chart
<!-- TODO: Add .gif for this -->
* Zoom-in/out a section of chart
<!-- TODO: Add .gif for this -->
* Save the png file
<!-- TODO: Add .gif for this -->
<div style="page-break-after: always;"></div>

## Troubleshooting

<div style="page-break-after: always;"></div>
