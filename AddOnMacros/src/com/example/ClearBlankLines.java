/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.example;

import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XLineCursor;
import com.sun.star.view.XViewCursor;

/**
 *
 * @author brant
 */
public class ClearBlankLines{

    private XTextViewCursor textViewCursor;
    private XViewCursor viewCursor;
    private XLineCursor lineCursor;

    public ClearBlankLines(XComponentContext xComponentContext) throws Exception {

    }
}
