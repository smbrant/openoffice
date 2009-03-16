/*
 * Macro3.java
 *
 * Created on 2009.01.29 - 12:35:57
 *
 * Creates a new calc document
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
public class Macro3 {
    
    /** Creates a new instance of Macro3 */
    public Macro3() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            // get the remote office component context
            XComponentContext context = Bootstrap.bootstrap();

            XMultiComponentFactory srvMan = context.getServiceManager();

//            String[] services= srvMan.getAvailableServiceNames();
//            for (String service : services)
//                System.out.println(service);
//
//            System.out.println("Total de serviços disponíveis: "+services.length);

            XDesktop desktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));


// Poderia pegar o desktop de forma mais simples, perdendo a  interface XDesktop:
//
//         Object desktop = srvMan.createInstanceWithContext("com.sun.star.frame.Desktop", context );
//
//

            XComponentLoader componentLoader =
                (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);

            PropertyValue dummy[] = new PropertyValue[0];



            XComponent doc= componentLoader.loadComponentFromURL(
                "private:factory/scalc","_blank", 0, dummy);

            

        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
