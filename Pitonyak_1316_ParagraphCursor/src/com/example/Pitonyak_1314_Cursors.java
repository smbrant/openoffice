/*
 * Pitonyak_1314_Cursors.java
 *
 * Created on 2009.03.16 - 14:20:15
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_1314_Cursors {
    
    /** Creates a new instance of Pitonyak_1314_Cursors */
    public Pitonyak_1314_Cursors() {
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

            XComponent document = (XComponent)
                UnoRuntime.queryInterface(XComponent.class, desktop.getCurrentComponent());

            XTextDocument textDocument = (XTextDocument) UnoRuntime.queryInterface ( XTextDocument.class,
                       document);

            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, document);

            if (model==null){
                System.out.println("Open a document please.");
            }else{
                XTextViewCursorSupplier viewCursorSupplier = (XTextViewCursorSupplier)UnoRuntime.queryInterface(XTextViewCursorSupplier.class,
                            model.getCurrentController());

                XTextViewCursor textViewCursor = viewCursorSupplier.getViewCursor();
                textViewCursor.gotoStart(false); // start of text

                XText text = textDocument.getText();

                XTextCursor textCursor = text.createTextCursor();

                XParagraphCursor paragraphCursor = (XParagraphCursor)
                    UnoRuntime.queryInterface( XParagraphCursor.class, textCursor );
                do {
                    paragraphCursor.gotoEndOfParagraph(true);
                    System.out.println(paragraphCursor.getString()); //process paragraph
                } while (paragraphCursor.gotoNextParagraph(false));
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
