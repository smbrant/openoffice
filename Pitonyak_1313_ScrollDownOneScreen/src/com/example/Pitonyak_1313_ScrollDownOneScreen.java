/*
 * Pitonyak_1313_ScrollDownOneScreen.java
 *
 * Created on 2009.02.12 - 11:11:45
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
import com.sun.star.view.XLineCursor;
import com.sun.star.view.XScreenCursor;
import com.sun.star.view.XViewCursor;

public class Pitonyak_1313_ScrollDownOneScreen {
    
    /** Creates a new instance of Pitonyak_1313_ScrollDownOneScreen */
    public Pitonyak_1313_ScrollDownOneScreen() {
    }
    
    //Sub ScrollDownOneScreen
    //    REM Get the view cursor from the current controller
    //    ThisComponent.currentController.getViewCursor().screenDown()
    //End Sub
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

                // The TextViewCursor have some operations:
                XTextViewCursor textViewCursor = viewCursorSupplier.getViewCursor();
                textViewCursor.gotoStart(false); // start of text

                // For more sophisticated operations use other interfaces to viewcursor:
                XScreenCursor screenCursor= (XScreenCursor) UnoRuntime.queryInterface(XScreenCursor.class,
                        textViewCursor);
                screenCursor.screenDown(); //pagedown

                //and one more line
                XViewCursor viewCursor= (XViewCursor) UnoRuntime.queryInterface(XViewCursor.class,
                        textViewCursor);
                viewCursor.goDown((short)1, false); // one line more

                //and select the last line
                XLineCursor lineCursor= (XLineCursor) UnoRuntime.queryInterface(XLineCursor.class,
                        viewCursor);
                lineCursor.gotoEndOfLine(true); // select until end of current line
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
