/*
 * From http://user.services.openoffice.org/en/forum/viewtopic.php?f=45&t=3813&p=17088&hilit=dialog#p17088
 * 
 */

package com.example;


import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
* The source code of this class bases heavily on the OOo Developer's Guide:
* http://api.openoffice.org/docs/DevelopersGuide/GUI/GUI.xhtml#1_Graphical_User_Interfaces
*
* This class provides methods to create, execute and dispose an OOo Dialog
* with labels (fixed text), buttons, text fields, and password fields.
*/
public class Dialog {

    private XComponentContext xComponentContext;
    private XMultiComponentFactory xMultiComponentFactory;
    private XMultiServiceFactory xMultiComponentFactoryDialogModel;
    private XNameContainer xDialogModelNameContainer;
    private XControl xDialogControl;
    private XControlContainer xDialogContainer;

    public Dialog(XComponentContext xComponentContext, int posX, int posY, int height, int width, String title, String name) throws Exception {

        this.xComponentContext = xComponentContext;
        init(posX, posY, height, width, title, name);
    }

    private void init(int posX, int posY, int height, int width, String title, String name) throws Exception {

        xMultiComponentFactory = xComponentContext.getServiceManager();

        Object oDialogModel =  xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", xComponentContext);

        // The XMultiServiceFactory of the dialogmodel is needed to instantiate the controls...
        xMultiComponentFactoryDialogModel = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, oDialogModel);

        // The named container is used to insert the created controls into...
        xDialogModelNameContainer = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, oDialogModel);

        // create the dialog...
        Object oUnoDialog = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", xComponentContext);
        xDialogControl = (XControl) UnoRuntime.queryInterface(XControl.class, oUnoDialog);

        // The scope of the control container is public...
        xDialogContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, oUnoDialog);

        // link the dialog and its model...
        XControlModel xControlModel = (XControlModel) UnoRuntime.queryInterface(XControlModel.class, oDialogModel);
        xDialogControl.setModel(xControlModel);

        // Create a unique name
        String uniqueName = createUniqueName(xDialogModelNameContainer, (name != null)? name: "OOoDialog");

        // Define the dialog at the model
        XPropertySet xDialogPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xDialogModelNameContainer);
        xDialogPropertySet.setPropertyValue("PositionX",new Integer(posX));
        xDialogPropertySet.setPropertyValue("PositionY",new Integer(posY));
        xDialogPropertySet.setPropertyValue("Height",new Integer(height));
        xDialogPropertySet.setPropertyValue("Width",new Integer(width));
        xDialogPropertySet.setPropertyValue("Title",(title != null)? title: "OpenOffice.org Dialog");
        xDialogPropertySet.setPropertyValue("Name",uniqueName);
        xDialogPropertySet.setPropertyValue("Moveable",Boolean.TRUE);
        xDialogPropertySet.setPropertyValue("TabIndex",new Short((short) 0));
    }

    public short execute() throws Exception {

        XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xDialogContainer);

        // set the dialog invisible until it is executed
        xWindow.setVisible(false);

        Object oToolkit = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.Toolkit", xComponentContext);
        XWindowPeer xWindowParentPeer = ((XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit)).getDesktopWindow();
        XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);
        xDialogControl.createPeer(xToolkit, xWindowParentPeer);

        // the return value contains information about how the dialog has been closed
        XDialog xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, xDialogControl);
        return xDialog.execute();
    }

    public void dispose() {

        // Free the resources
        XComponent xDialogComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xDialogControl);
        xDialogComponent.dispose();
    }

    public XFixedText insertFixedText(int posX, int posY, int width, String label) throws Exception {

        // Create a unique name
        String uniqueName = createUniqueName(xDialogModelNameContainer, "FixedText");

        // Create a fixed text control model
        Object oFixedTextModel = xMultiComponentFactoryDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel");

        // Set the properties at the model
        XPropertySet xFixedTextPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oFixedTextModel);
        xFixedTextPropertySet.setPropertyValue("PositionX",new Integer(posX));
        xFixedTextPropertySet.setPropertyValue("PositionY",new Integer(posY+2));
        xFixedTextPropertySet.setPropertyValue("Height",new Integer(8));
        xFixedTextPropertySet.setPropertyValue("Width",new Integer(width));
        xFixedTextPropertySet.setPropertyValue("Label",(label != null)? label: "");
        xFixedTextPropertySet.setPropertyValue("Name",uniqueName);

        // Add the model to the dialog model name container
        xDialogModelNameContainer.insertByName(uniqueName, oFixedTextModel);

        // Reference the control by the unique name
        XControl xFixedTextControl = xDialogContainer.getControl(uniqueName);

        return (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xFixedTextControl);
    }

    public XButton insertButton(int posX, int posY, int width, String label, int pushButtonType) throws Exception {

        // Create a unique name
        String uniqueName = createUniqueName(xDialogModelNameContainer, "Button");

        // Create a button control model
        Object oButtonModel = xMultiComponentFactoryDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel");

        // Set the properties at the model
        XPropertySet xButtonPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oButtonModel);
        xButtonPropertySet.setPropertyValue("PositionX",new Integer(posX));
        xButtonPropertySet.setPropertyValue("PositionY",new Integer(posY));
        xButtonPropertySet.setPropertyValue("Height",new Integer(14));
        xButtonPropertySet.setPropertyValue("Width",new Integer(width));
        xButtonPropertySet.setPropertyValue("Label",(label != null)? label: "");
        xButtonPropertySet.setPropertyValue("PushButtonType",new Short((short) pushButtonType));
        xButtonPropertySet.setPropertyValue("Name",uniqueName);

        // Add the model to the dialog model name container
        xDialogModelNameContainer.insertByName(uniqueName, oButtonModel);

        // Reference the control by the unique name
        XControl xButtonControl = xDialogContainer.getControl(uniqueName);

        return (XButton) UnoRuntime.queryInterface(XButton.class, xButtonControl);
    }

    public XTextComponent insertTextField(int posX, int posY, int width, String text) throws Exception {

        return insertEditableTextField(posX, posY, width, text, ' ');
    }

    public XTextComponent insertPasswordField(int posX, int posY, int width, String text, char echoChar) throws Exception {

        return insertEditableTextField(posX, posY, width, text, echoChar);
    }

    private XTextComponent insertEditableTextField(int posX, int posY, int width, String text, char echoChar) throws Exception {

        // Create a unique name
        String uniqueName = createUniqueName(xDialogModelNameContainer, "EditableTextField");

        // Create an editable text field control model
        Object oEditableTextFieldModel = xMultiComponentFactoryDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel");

        // Set the properties at the model
        XPropertySet xEditableTextFieldPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oEditableTextFieldModel);
        xEditableTextFieldPropertySet.setPropertyValue("PositionX",new Integer(posX));
        xEditableTextFieldPropertySet.setPropertyValue("PositionY",new Integer(posY));
        xEditableTextFieldPropertySet.setPropertyValue("Height",new Integer(12));
        xEditableTextFieldPropertySet.setPropertyValue("Width",new Integer(width));
        xEditableTextFieldPropertySet.setPropertyValue("Text",(text != null)? text: "");
        xEditableTextFieldPropertySet.setPropertyValue("Name",uniqueName);

        if (echoChar != 0 && echoChar != ' ') {
            // Useful for password fields
            xEditableTextFieldPropertySet.setPropertyValue("EchoChar",new Short((short) echoChar));
        }

        // Add the model to the dialog model name container
        xDialogModelNameContainer.insertByName(uniqueName, oEditableTextFieldModel);

        // Reference the control by the unique name
        XControl xEditableTextFieldControl = xDialogContainer.getControl(uniqueName);

        return (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, xEditableTextFieldControl);
    }

    /**
     * Makes a string unique by appending a numerical suffix.
     *
     * @param elementContainer   The container the new element is going to be inserted to
     * @param elementName        The name of the element
     */
    private static String createUniqueName(XNameAccess elementContainer, String elementName) {

        String uniqueElementName = elementName;

        boolean elementExists = true;
        int i = 1;
        while (elementExists) {
            elementExists = elementContainer.hasByName(uniqueElementName);
            if (elementExists) {
                i++;
                uniqueElementName = elementName + Integer.toString(i);
            }
        }

        return uniqueElementName;
    }
}