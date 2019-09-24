using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace UClass.Offic.PPT
{
    public class PreViewClass
    {
        public event Action<int> SlideShowNextSlide;
        public int SlidesCount;
        public int HWND;

        private Microsoft.Office.Interop.PowerPoint.Application oPPT;
        public void Start(string pptFileFullName)
        {
            Application objApp = Marshal.GetActiveObject("PowerPoint.Application") as Application; //new Application();
            //var objPresSet = objApp.Presentations;
            objApp.SlideShowNextSlide += ObjApp_SlideShowNextSlide;
            Presentation objPres = objApp.ActivePresentation;// objPresSet.Open(pptFileFullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);




            var objSlides = objPres.Slides;
            SlidesCount = objSlides.Count;

            int[] SlideIdx = new int[3];
            for (int i = 0; i < 3; i++)
                SlideIdx[i] = i + 1;
            var objSldRng = objSlides.Range(SlideIdx);
            var objSST = objSldRng.SlideShowTransition;
            objSST.AdvanceOnTime = MsoTriState.msoTrue; objSST.AdvanceTime = 3;
            objSST.EntryEffect = Microsoft.Office.Interop.PowerPoint.PpEntryEffect.ppEffectBoxOut;

            var objSSS = objPres.SlideShowSettings;
            //如过你不想循环放映就把TRUE改成FALSE. 
            objSSS.LoopUntilStopped = MsoTriState.msoFalse;
            objSSS.StartingSlide = 1;

            objSSS.EndingSlide = objSlides.Application.ActivePresentation.Slides.Count;
            objSSS.Run(); //Wait for the slide show to end. 

            //翻到指定页
            //objPres.SlideShowWindow.View.GotoSlide(16);
        }

        private void ObjApp_SlideShowNextSlide(SlideShowWindow Wn)
        {
            int p = Wn.View.CurrentShowPosition;
            SlideShowNextSlide?.Invoke(p);

            HWND = Wn.HWND;
            
        }
    }
}
