/*
 * Macro_17_SaveAs.java
 *
 * Created on 2009.03.12 - 13:20:36
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_17_SaveAs {
    
    /** Creates a new instance of Macro_17_SaveAs */
    public Macro_17_SaveAs() {
    }


    public static void main(String[] args) throws Exception {


        XComponentContext context = Bootstrap.bootstrap();
        XMultiComponentFactory serviceManager = context.getServiceManager();
        XDesktop desktop =
          (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
              serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
        XComponent document= desktop.getCurrentComponent();

        if (document != null){
            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, document);

            XController controller= model.getCurrentController();
            XFrame frame=controller.getFrame();

            XDispatchProvider dispatchProvider = (XDispatchProvider)UnoRuntime.queryInterface(XDispatchProvider.class, frame);

            XDispatchHelper dispatchHelper =
              (XDispatchHelper) UnoRuntime.queryInterface( XDispatchHelper.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper",context));

            dispatchHelper.executeDispatch(dispatchProvider, ".uno:SaveAs", "", 0, null);
        }else{
            System.out.println("No open document. Nothing to do...");
        }

        System.exit( 0 );

//   oDoc = ThisComponent
//   oDocCtrl = oDoc.getCurrentController()
//   oDocFrame = oDocCtrl.getFrame()
//
//   oDispatchHelper = createUnoService("com.sun.star.frame.DispatchHelper")
//   oDispatchHelper.executeDispatch(oDocFrame, ".uno:SaveAs", "", 0, Array())
    }


    
}
