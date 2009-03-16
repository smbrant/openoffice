/*
 * Pitonyak_1314_InsertSpecialCharacter.java
 *
 * Created on 2009.03.16 - 12:39:44
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_1314_InsertSpecialCharacter {
    
    /** Creates a new instance of Pitonyak_1314_InsertSpecialCharacter */
    public Pitonyak_1314_InsertSpecialCharacter() {
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
            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, document);

            if (model==null){
                System.out.println("Open a document please.");
            }else{
                XTextViewCursorSupplier viewCursorSupplier = (XTextViewCursorSupplier)UnoRuntime.queryInterface(XTextViewCursorSupplier.class,
                            model.getCurrentController());

                XTextViewCursor textViewCursor = viewCursorSupplier.getViewCursor();
                textViewCursor.getText().insertString(textViewCursor.getStart(),"\u0257" , false);

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
