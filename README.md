# StarCCM Automation
This macro was designed to hep automate your STAR-CCM post-processes. This repo contains a macro which exports all scenes and plots to a separate folder. Once completed, run the supplied python file to complete post processing by generating a presentation with all scenes and plots labelled. Deeply useful for FSAE, Formula Student and other projects which require documentation of multiple simulations.
NOTE: You should be able to customize this project using the .py file.

Files included:
1. ExportScenesAndPlots.java: This is the inital macro to be used in STAR-CCM. This downloads the scenes and plots to an external folder.
2. NitroMacro.py: This is to be ran after the scenes and plots are exported. This generates the presentation.
3. NitroMacro.exe: Same thing, but an executable. should prevent python specific errors on different devices due to not importing modules.
4. Demo_Presentation.pptx: A demo presentation, shows what the current presentation generated should look like.

STEPS TO RUN:
1. Go to your active simulation file in STAR-CCM. Hit File --> Macro --> Play Macro --> Select ExportScenesAndPlots.java
2. Wait for the macro to finish. It will generate the folder "(Sim Name)_results"
3. Once macro is completed, run the python file/ executable depending on preference.
4. This will open a file picker. Pick your simulation results FOLDER.
5. The final presentation should be generated and exported into the same initial folder, "(Sim Name)_results"

PATCH NOTES:

V 1.0 -->
- Created base program
- Generates plots and scenes into ppt
  
V 1.1 --> 
- Upgraded base program.
- removed scaling and placement issues.
- Added title page with simulation name and company name.
- All images are now auto labelled with image names.
- Erased blank pages and removed placeholders obstructing images.
- Overall stability upgrades from V 1.0.


