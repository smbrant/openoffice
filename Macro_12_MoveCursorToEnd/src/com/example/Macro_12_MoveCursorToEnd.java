/*
 * Macro_12_MoveCursorToEnd.java
 *
 * Created on 2009.02.13 - 09:33:21
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_12_MoveCursorToEnd {
    
    /** Creates a new instance of Macro_12_MoveCursorToEnd */
    public Macro_12_MoveCursorToEnd() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();

            XDesktop desktop= (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XTextDocument textDocument = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, desktop.getCurrentComponent());

            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, textDocument);

            XTextViewCursorSupplier viewCursorSupplier = (XTextViewCursorSupplier)UnoRuntime.queryInterface(XTextViewCursorSupplier.class,
                        model.getCurrentController());

            XTextViewCursor viewCursor = viewCursorSupplier.getViewCursor();
            viewCursor.gotoEnd(true);
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
