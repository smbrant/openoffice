/*
 * Macro_07.java
 *
 * Created on 2009.02.04 - 15:34:03
 *
 * Creating a dialog
 *
 * From http://www.oooforum.org/forum/viewtopic.phtml?t=10023
 *
 */

package com.example;

import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

public class Macro_07 {
    static private XComponentContext context;
    static private XMultiComponentFactory serviceMan;
    
    /** Creates a new instance of Macro_07 */
    public Macro_07() {
    }

    public static void main(String[] args) {
        try {
            context = Bootstrap.bootstrap();
            serviceMan = context.getServiceManager();
            createDialog();
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
    /** method for creating a dialog at runtime
     */
    static void createDialog() throws com.sun.star.uno.Exception {
         // create the dialog model and set the properties
         Object dialogModel = serviceMan.createInstanceWithContext(
             "com.sun.star.awt.UnoControlDialogModel", context);
         XPropertySet xPSetDialog = (XPropertySet)UnoRuntime.queryInterface(
             XPropertySet.class, dialogModel);
         xPSetDialog.setPropertyValue("PositionX", new Integer(100));
         xPSetDialog.setPropertyValue("PositionY", new Integer(100));
         xPSetDialog.setPropertyValue("Width", new Integer(150));
         xPSetDialog.setPropertyValue("Height", new Integer(100));
         xPSetDialog.setPropertyValue("Title", new String("Runtime Dialog Demo"));

         // get the service manager from the dialog model
         XMultiServiceFactory xMultiServiceFactory = (XMultiServiceFactory)UnoRuntime.queryInterface(
             XMultiServiceFactory.class, dialogModel);

         // create the button model and set the properties
         Object buttonModel = xMultiServiceFactory.createInstance(
             "com.sun.star.awt.UnoControlButtonModel" );
         XPropertySet xPSetButton = (XPropertySet)UnoRuntime.queryInterface(
             XPropertySet.class, buttonModel);
         xPSetButton.setPropertyValue("PositionX", new Integer(50));
         xPSetButton.setPropertyValue("PositionY", new Integer(30));
         xPSetButton.setPropertyValue("Width", new Integer(50));
         xPSetButton.setPropertyValue("Height", new Integer(14));
         xPSetButton.setPropertyValue("Name", new String("_buttonName"));
         xPSetButton.setPropertyValue("TabIndex", new Short((short)0));
         xPSetButton.setPropertyValue("Label", new String("Click Me"));

         // create the label model and set the properties
         Object labelModel = xMultiServiceFactory.createInstance(
             "com.sun.star.awt.UnoControlFixedTextModel" );
         XPropertySet xPSetLabel = ( XPropertySet )UnoRuntime.queryInterface(
             XPropertySet.class, labelModel );
         xPSetLabel.setPropertyValue("PositionX", new Integer(40));
         xPSetLabel.setPropertyValue("PositionY", new Integer(60));
         xPSetLabel.setPropertyValue("Width", new Integer(100));
         xPSetLabel.setPropertyValue("Height", new Integer(14));
         xPSetLabel.setPropertyValue("Name", new String("_labelName"));
         xPSetLabel.setPropertyValue("TabIndex", new Short((short)1));
         xPSetLabel.setPropertyValue("Label", new String("_labelPrefix"));

         // insert the control models into the dialog model
         XNameContainer xNameCont = (XNameContainer)UnoRuntime.queryInterface(
             XNameContainer.class, dialogModel);
         xNameCont.insertByName("_buttonName", buttonModel);
         xNameCont.insertByName("_labelName", labelModel);

         // create the dialog control and set the model
         Object dialog = serviceMan.createInstanceWithContext(
             "com.sun.star.awt.UnoControlDialog", context);
         XControl xControl = (XControl)UnoRuntime.queryInterface(
             XControl.class, dialog );
         XControlModel xControlModel = (XControlModel)UnoRuntime.queryInterface(
             XControlModel.class, dialogModel);
         xControl.setModel(xControlModel);

         // add an action listener to the button control
         XControlContainer xControlCont = (XControlContainer)UnoRuntime.queryInterface(
             XControlContainer.class, dialog);
         Object objectButton = xControlCont.getControl("Button1");
         XButton xButton = (XButton)UnoRuntime.queryInterface(XButton.class, objectButton);
//         xButton.addActionListener(new ActionListenerImpl(xControlCont));

         // create a peer
         Object toolkit = serviceMan.createInstanceWithContext(
             "com.sun.star.awt.Toolkit", context);
         XToolkit xToolkit = (XToolkit)UnoRuntime.queryInterface(XToolkit.class, toolkit);
         XWindow xWindow = (XWindow)UnoRuntime.queryInterface(XWindow.class, xControl);
         xWindow.setVisible(false);
         xControl.createPeer(xToolkit, null);

         // execute the dialog
         XDialog xDialog = (XDialog)UnoRuntime.queryInterface(XDialog.class, dialog);
         xDialog.execute();

         // dispose the dialog
         XComponent xComponent = (XComponent)UnoRuntime.queryInterface(XComponent.class, dialog);
         xComponent.dispose();
     }

}
