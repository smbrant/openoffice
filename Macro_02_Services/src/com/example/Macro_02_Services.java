/*
 * Macro_02_Services.java
 *
 * Created on 2009.03.03 - 12:06:09
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
public class Macro_02_Services {
    
    /** Creates a new instance of Macro_02_Services */
    public Macro_02_Services() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            // get the remote office component context
            XComponentContext context = Bootstrap.bootstrap();

            XMultiComponentFactory srvMan = context.getServiceManager();

            System.out.println("Services:");
            String[] services= srvMan.getAvailableServiceNames();
            for (String service : services)
                System.out.println(service);

            System.out.println("Total available services: "+services.length);

        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
