/*
 * Macro_15_SimpleLoader.java
 *
 * Created on 2009.03.04 - 09:43:47
 *
 */

package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_15_SimpleLoader {
    
    /** Creates a new instance of Macro_15_SimpleLoader */
    public Macro_15_SimpleLoader() {
    }
    
    public static void main(String[] args) {
        try {
            XComponentContext xContext = Bootstrap.bootstrap();

            XMultiComponentFactory xMCF = xContext.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",xContext));

            XComponentLoader xComponentLoader =
                (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);

            PropertyValue xEmptyArgs[] = new PropertyValue[0];

            XComponent xComponent =xComponentLoader.loadComponentFromURL(
                "private:factory/swriter","_blank", 0, xEmptyArgs);

            XTextDocument xTextDocument =
                (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComponent);

            XText xText = xTextDocument.getText();

            XTextCursor xTextCursor = (XTextCursor) xText.createTextCursor();

            xText.insertString( xTextCursor, "Hello from a simple loader", false );

        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
