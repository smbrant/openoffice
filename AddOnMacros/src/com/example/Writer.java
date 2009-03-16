/*
 * Interface do sct com o brOffice.
 *
 */

package com.example;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XLineCursor;
import com.sun.star.view.XViewCursor;

/**
 * Writer: encapsula macros do sct
 *
 * @author brant
 * 
 */
public class Writer {

    private XMultiComponentFactory serviceManager;
    private XDesktop desktop;
    private XComponent document;
    private XModel model;
    private XTextViewCursor textViewCursor;
    private XViewCursor viewCursor;
    private XLineCursor lineCursor;

    public Writer(XComponentContext context) throws Exception {
        serviceManager = context.getServiceManager();
        desktop= (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
            serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
        document = (XComponent) UnoRuntime.queryInterface(XComponent.class,
            desktop.getCurrentComponent());
        model= (XModel) UnoRuntime.queryInterface(XModel.class, document);
    }

    private void clearLineIfBlank(){
        lineCursor.gotoStartOfLine(false);
        lineCursor.gotoEndOfLine(true);
        while (textViewCursor.getString().equals("") && viewCursor.goDown((short)1, true)){
            lineCursor.gotoStartOfLine(true);
            textViewCursor.setString("");
            lineCursor.gotoStartOfLine(false);
            lineCursor.gotoEndOfLine(true);
        }
    }

    private void clearLastLineIfBlank(){
        lineCursor.gotoStartOfLine(false);
        lineCursor.gotoEndOfLine(true);
        if (textViewCursor.getString().equals("") && viewCursor.goUp((short)1, true)){
            lineCursor.gotoEndOfLine(true);
            textViewCursor.setString("");
        }
    }

    public void clearBlankLines(){
        if (model!=null){
            XTextViewCursorSupplier viewCursorSupplier = (XTextViewCursorSupplier)UnoRuntime.queryInterface(XTextViewCursorSupplier.class,
                        model.getCurrentController());
            textViewCursor = viewCursorSupplier.getViewCursor();
            viewCursor= (XViewCursor) UnoRuntime.queryInterface(XViewCursor.class, textViewCursor);
            lineCursor= (XLineCursor) UnoRuntime.queryInterface(XLineCursor.class, viewCursor);

            textViewCursor.gotoStart(false); // start of text
            clearLineIfBlank();
            while (viewCursor.goDown((short)1, false)){
                clearLineIfBlank();
            }
            clearLastLineIfBlank();
            //TODO: se o cursor estiver dentro do cabeçalho a macro não funciona.
        }
    }

    public void clearHeaderFooter(){
        XTextDocument textDocument = (XTextDocument)
            UnoRuntime.queryInterface(XTextDocument.class, desktop.getCurrentComponent());
        if (textDocument!=null)
            try{
                XText text = textDocument.getText();
                XTextCursor textCursor = text.createTextCursor();
                XPropertySet cursorProperties= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, textCursor);
                String pageStyleName= (String) cursorProperties.getPropertyValue("PageStyleName");
                XStyleFamiliesSupplier styleFamiliesSupplier= (XStyleFamiliesSupplier) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, textDocument);
                XNameAccess stylesFamilies= styleFamiliesSupplier.getStyleFamilies();

                XNameAccess styleFamily= (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class,
                        stylesFamilies.getByName("PageStyles"));
                XPropertySet pageProperties= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
                        styleFamily.getByName(pageStyleName) );

                if((Boolean)pageProperties.getPropertyValue("HeaderIsOn")){
                    pageProperties.setPropertyValue("HeaderIsOn", new Boolean(false));
                }
                if((Boolean)pageProperties.getPropertyValue("FooterIsOn")){
                    pageProperties.setPropertyValue("FooterIsOn", new Boolean(false));
                }
            }
            catch (java.lang.Exception e){e.printStackTrace();}
            finally {return;}
    }
}
