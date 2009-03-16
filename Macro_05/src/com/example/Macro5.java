/*
 * Macro5.java
 *
 * Created on 2009.02.02 - 15:02:14
 *
 */
        /*
         * This snippet shows how to create a frame at an arbitrary location and ZOrder within a table.
         * The Parameter IsFollowingTextFlow ensures that the table "grows"
         * The XText does not neccessarily have to be from a table.
         * It also shows how to make the frame transparent
         *
         */

package com.example;

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro5 {
    
    /** Creates a new instance of Macro5 */
    public Macro5() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
        /* This snippet shows how to create a frame at an arbitrary location and ZOrder within a table.
         * The Parameter IsFollowingTextFlow ensures that the
         table "grows"
         * The XText does not neccessarily have to be from a table.
         * It also shows how to make the frame transparent
         *
         * */

        XMultiServiceFactory documentFactory = null;
        XText currentXCellText = null;
        Integer x = new Integer( 20 );
        Integer y = new Integer( 20 );

        int width = 50;
        int height = 50;

        Integer ZOrder = new Integer( 1 );

        try
        {
          Object writerShape = documentFactory.createInstance( "com.sun.star.text.TextFrame" );

          XShape xWriterShape = ( XShape ) UnoRuntime.queryInterface( XShape.class, writerShape );
          xWriterShape.setSize( new Size( width, height ) );

          XTextContent xTextContentShape = ( XTextContent ) UnoRuntime.queryInterface( XTextContent.class, writerShape );

          // does not support XFastPropertySet
          XPropertySet xTextContentPropertySet = ( XPropertySet ) UnoRuntime.queryInterface( XPropertySet.class,
                                                                                             xTextContentShape );
          xTextContentPropertySet.setPropertyValue( "FrameStyleName", "FrameStyle" );
          xTextContentPropertySet.setPropertyValue( "FrameIsAutomaticHeight", Boolean.TRUE );
          xTextContentPropertySet.setPropertyValue( "ZOrder", ZOrder );
          xTextContentPropertySet.setPropertyValue( "IsFollowingTextFlow", Boolean.TRUE );
          xTextContentPropertySet.setPropertyValue( "BackColor", new Integer( 0xffffffff ) );
          //$NON-NLS-1$
          xTextContentPropertySet.setPropertyValue( "BackColorTransparency", new Short( ( short ) 100 ) );
          //$NON-NLS-1$

          XPropertySet xShapeProps = ( XPropertySet ) UnoRuntime.queryInterface( XPropertySet.class, writerShape );

          // Setting the vertical position
          xShapeProps.setPropertyValue( "HoriOrientPosition", x );
          xShapeProps.setPropertyValue( "VertOrientPosition", y );

          // get the XText from the shape
          XText xShapeText = ( XText ) UnoRuntime.queryInterface( XText.class, writerShape );

          //TODO: não sei o que é o this neste exemplo
//          currentXCellText.insertTextContent( this.currentXCellCursor, xTextContentShape, false );

          xShapeText.setString( "SOME TEXT " );
        }
        catch ( Exception e )
        {
          e.printStackTrace();
        }
        System.exit( 0 );
    }
    
}
