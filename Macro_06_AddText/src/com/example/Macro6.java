/*
 * Macro6.java
 *
 * Created on 2009.02.03 - 10:55:01
 *
 * Adds text to the current text document.
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro6 {
    
    /** Creates a new instance of Macro6 */
    public Macro6() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            XComponentContext xContext = Bootstrap.bootstrap();

            XMultiComponentFactory xMCF = xContext.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",xContext));

            XTextDocument xtextdoc = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());
            if (xtextdoc!=null){
                XText xtext = xtextdoc.getText();
                XTextCursor xtextcursor = xtext.createTextCursor();
                xtext.insertString(xtextcursor, "Insert text if is a text document", false);
            }else{
                System.out.println("You should have an open text document.");
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
