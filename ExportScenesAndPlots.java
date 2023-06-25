import star.common.*;
import star.vis.*;
import java.io.File;
import java.util.Scanner; 

public class ExportScenesAndPlots extends StarMacro {
    public void execute() {

        Simulation sim = getActiveSimulation();
        
        String dir = sim.getSessionDir(); //get the name of the simulation's directory
        String sep = System.getProperty("file.separator"); //get the right separator for your operative system
	  File resultFolder = new File(dir + "\\" + sim + "_Results");
	  resultFolder.mkdir();

	  File sceneFolder = new File(resultFolder + "\\Scenes");
	  sceneFolder.mkdir();

	  File plotFolder = new File(resultFolder + "\\Plots");
	  plotFolder.mkdir();
              

        for (Scene scn: sim.getSceneManager().getScenes()) {
            sim.println("Saving Scene: " + scn.getPresentationName());
	      scn.printAndWait(resolvePath(sceneFolder + sep + scn.getPresentationName() + ".jpg"), 1, 1920, 1080);
		
	   	
        }

	  for (StarPlot plt : sim.getPlotManager().getObjects()) {
	      sim.println("Saving Plot: " + plt.getPresentationName());
	      plt.encode(resolvePath(plotFolder + sep + plt.getPresentationName() + ".jpg"), "jpg", 2560, 1440);
	  }
    }
}