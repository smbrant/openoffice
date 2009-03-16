/*
 * Macro4.java
 *
 * Created on 2009.02.02 - 14:56:34
 *
 * Macro_03 simplified
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
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro4 {
    
    /** Creates a new instance of Macro4 */
    public Macro4() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
        XComponentContext context = Bootstrap.bootstrap();

        XMultiComponentFactory srvMan = context.getServiceManager();

        XDesktop desktop =
          (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
              srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));

        XComponentLoader componentLoader =
            (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);

        PropertyValue dummy[] = new PropertyValue[0];

        XComponent doc= componentLoader.loadComponentFromURL(
            "private:factory/scalc","_blank", 0, dummy);
        System.exit( 0 );
    }
    
}
