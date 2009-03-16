/*
 * Macro_13_ClearBlankLines.java
 *
 * Created on 2009.02.13 - 09:35:34
 *
 * Clear blank lines
 *
 */

package com.example;

import com.sun.star.awt.XTextComponent;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.view.XLineCursor;
import com.sun.star.view.XViewCursor;

/**
 *
 * @author brant
 */
public class Macro_13_ClearBlankLines {
    
    /** Creates a new instance of Macro_13_ClearBlankLines */
    public Macro_13_ClearBlankLines() {
    }

    static private XTextViewCursor textViewCursor;
    static private XViewCursor viewCursor;
    static private XLineCursor lineCursor;
    
    static private void clearLineIfBlank(){
        lineCursor.gotoStartOfLine(false);
        lineCursor.gotoEndOfLine(true);
        while (textViewCursor.getString().equals("") && viewCursor.goDown((short)1, true)){
            lineCursor.gotoStartOfLine(true);
            textViewCursor.setString("");
            lineCursor.gotoStartOfLine(false);
            lineCursor.gotoEndOfLine(true);
        }
    }

    static private void clearLastLineIfBlank(){
        lineCursor.gotoStartOfLine(false);
        lineCursor.gotoEndOfLine(true);
        if (textViewCursor.getString().equals("") && viewCursor.goUp((short)1, true)){
            lineCursor.gotoEndOfLine(true);
            textViewCursor.setString("");
        }
    }

    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();

            XDesktop desktop= (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent document = (XComponent)
                UnoRuntime.queryInterface(XComponent.class, desktop.getCurrentComponent());
            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, document);

            if (model==null){
                System.out.println("Open a document please.");
            }else{
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
            }
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
