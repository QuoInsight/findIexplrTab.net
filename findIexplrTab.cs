using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Accessibility;
using System.Text;
using System.Collections.Generic;
using System.Threading;

namespace myNameSpace {
  class myClass {

  /*
    [DllImport("oleacc.dll")]
    static extern uint WindowFromAccessibleObject(IAccessible pacc, ref IntPtr phwnd);
  */

    [DllImport("oleacc.dll")]
    static extern int AccessibleObjectFromWindow(
      IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject
    );

    static void getAccessibleObjectFromWindow(IntPtr hwnd, ref Accessibility.IAccessible acc) {
      Guid guid = new Guid("{618736e0-3c3d-11cf-810c-00aa00389b71}"); // IAccessibleobject obj = null;
      object tmpObj = null;
      int returnVal = AccessibleObjectFromWindow(hwnd, (uint)0x00000000, ref guid, ref tmpObj);
      acc = (Accessibility.IAccessible) tmpObj;
      return;
    }

    [DllImport("oleacc.dll")]
    static extern int AccessibleChildren(
      Accessibility.IAccessible paccContainer, int iChildStart, int cChildren,
      [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] rgvarChildren,
      out int pcObtained
    );

    static IAccessible[] getAccessibleChildren(IAccessible objAccessible) {
      if (objAccessible.accChildCount < 1) return new IAccessible[0];

      int accChildCount=objAccessible.accChildCount, resultCount=0;
      object[] accChildren = new object[accChildCount];
      AccessibleChildren(objAccessible, 0, accChildCount, accChildren, out resultCount);

      var list = new List<IAccessible>(accChildren.Length);
      foreach (object obj in accChildren)  {
        var acc = obj as IAccessible;
        if (acc != null) list.Add(acc);
      }
      return list.ToArray();
    }

    [DllImport("User32.dll", SetLastError=true, CharSet=CharSet.Auto)] 
    static extern long GetClassName(IntPtr hwnd, StringBuilder lpClassName, long nMaxCount); 

    [DllImport("user32.dll", SetLastError=true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    static IntPtr GetDirectUIHWND(IntPtr hWnd0) {
      IntPtr hWnd1 = FindWindowEx(hWnd0, IntPtr.Zero, "WorkerW", null);
      if (hWnd1!=IntPtr.Zero) hWnd1 = FindWindowEx(hWnd1, IntPtr.Zero, "ReBarWindow32", null);
      if (hWnd1!=IntPtr.Zero) hWnd1 = FindWindowEx(hWnd1, IntPtr.Zero, "TabBandClass", null);
      if (hWnd1!=IntPtr.Zero) hWnd1 = FindWindowEx(hWnd1, IntPtr.Zero, "DirectUIHWND", null);
      return hWnd1;
    }

    static void printIexplrTabInfo(System.Diagnostics.Process iexplrProc, IAccessible accTab, int tabIndex, int i) {
      Console.WriteLine(
        "[PID:{0}] [HWND:{1}] [{2}] [tabIndex:{3}] [i:{4}] {5}",
        iexplrProc.Id, iexplrProc.MainWindowHandle.ToInt32(), iexplrProc.MainWindowTitle,
        tabIndex, i, accTab.accDescription[0].Replace(Environment.NewLine, " | ")
      );
    }

    static int findIexplrTab(System.Diagnostics.Process iexplrProc, string findTxt, bool activateTab) {
      // [ https://stackoverflow.com/questions/3820228/setting-focus-to-already-opened-tab-of-internet-explorer-from-c-sharp-program-us ]
      IntPtr hWndDirectUI=GetDirectUIHWND(iexplrProc.MainWindowHandle);  if (hWndDirectUI==IntPtr.Zero) return int.MinValue;

      // [ https://msdn.microsoft.com/en-us/library/system.windows.forms.accessibleobject(v=vs.110).aspx ]
      // [ https://www.codeproject.com/Articles/38906/UI-Automation-Using-Microsoft-Active-Accessibility ]
      Accessibility.IAccessible objAccessible=null;  getAccessibleObjectFromWindow(hWndDirectUI, ref objAccessible);
      if (objAccessible==null) {
        Console.WriteLine("ERROR: getAccessibleObjectFromWindow()");
        return int.MinValue;
      }

      bool foundTab=false; int tabIndex=0;
      foreach (IAccessible accessor in getAccessibleChildren(objAccessible)) {
        foreach (IAccessible accChild in getAccessibleChildren(accessor)) {
          IAccessible[] accTabs = getAccessibleChildren(accChild);
          for (int i=0; i<accTabs.Length-1; i++) {
            tabIndex++;  IAccessible accTab = accTabs[i];
            if ( findTxt.Length==0 ) {
              if ((int)accTab.get_accState(0)==0x200002) { // 2097154
                printIexplrTabInfo(iexplrProc, accTab, tabIndex, i);
                Console.WriteLine("[activeTab:{0}]", tabIndex);
                return tabIndex;
              }
            } else if ( findTxt.Length > 0 && accTab.accDescription[0].Contains(findTxt) ) {
              foundTab = true;
              printIexplrTabInfo(iexplrProc, accTab, tabIndex, i);
              Console.WriteLine("[foundTab:{0}]", tabIndex);
              if (activateTab && (int)accTab.get_accState(0)!=0x200002) {
                accTab.accDoDefaultAction(0);  // 0==CHILDID_SELF ==> !!activate this tab!!
                System.Threading.Thread.Sleep(200);
                if ((int)accTab.get_accState(0)==0x200002) {
                  Console.WriteLine("[activateTab:success]");
                } else {
                  Console.WriteLine("[activateTab:failed]");
                }
              }
              return tabIndex;
            }
          } // accTabs
        } // accChild
      } // accessor

      return -tabIndex;
    }

    static void Main(string[] args) {
      string findByTxt = (args.Length > 0) ? args[0] : "";

      System.Diagnostics.Process[] procList = System.Diagnostics.Process.GetProcessesByName("iexplore");
      foreach (var iexplrProc in procList) {
        if (iexplrProc.MainWindowTitle.Length > 0) {
          System.Text.StringBuilder wndClassName = new StringBuilder(256);;
          GetClassName(iexplrProc.MainWindowHandle, wndClassName, wndClassName.Capacity);
          if ( wndClassName.ToString()=="IEFrame" ) {
            int tabIndex = findIexplrTab(iexplrProc, findByTxt, true);
            if (tabIndex > 0) return; // found
          }
        }
      }

    }  // Main

  }  // myClass
}  // myNameSpace
