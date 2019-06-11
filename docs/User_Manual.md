# AutoPlot

<p align="left">
  <img src="../icons/autoplot_ic_launcher.png" alt="AutoPlot Icon" width="" height="">
</p>

## Installation
Please follow the [Installation]("../Installation") folder.

## Examples
This has been implemented in different case studies like __Dry Etch__, __Wet Etch__, __Diffusion__ so far.
### Modules

## Execution
1. ### M-1: `button-mode`
	- Features:
		+ Customize the Plot with more buttons (simple, checkbox type) on the Excel sheet.
2. ### M-2: `shell-mode`
	```bat
	cmd /c python run.py
	```
	- Features:
		+ Here, the charts will be generated/updated in the current directory.
		+ In the `run.py` file, the charts can be customized by editing in the script, using these:
			* `auto_open= True` (by default, no need)
			* `auto_open= False`
	
3. ### M-3: `auto-mode`
	- Features:	
		+ `auto_run.sh` is added via `auto_run.bat` in __Task Scheduler__ to automate the process.
