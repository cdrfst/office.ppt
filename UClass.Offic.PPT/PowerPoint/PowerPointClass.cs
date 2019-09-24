using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PPt = Microsoft.Office.Interop.PowerPoint;

namespace UClass.Office.PowerPoint
{
    [Export(typeof(IPowerPoint))]
    public class PowerPointClass : IPowerPoint
    {
        public event Action SlideShowEnd;
        public event Action<int> PageChanged;
        private Microsoft.Office.Interop.PowerPoint.Application _pptApplication;
        private Microsoft.Office.Interop.PowerPoint.Presentation _presentation;
        private PPt.Slide slide;
        private PPt.Slides slides;
        private bool _isGetPptObjectCompleted;
        private bool _isDisposed;
        public bool IsDisposed
        {
            get { return _isDisposed; }
        }
        public int PageCount { get; private set; }
        public int CurrentPage
        {
            get
            {
                try
                {
                    return _presentation.SlideShowWindow.View.CurrentShowPosition;
                }
                catch (Exception)
                {
                    //ignore
                    return 1;
                }
            }
        }

        public bool IsGetPptObjectCompleted
        {
            get { return _isGetPptObjectCompleted; }
        }

        public PowerPointClass()
        {

        }

        public void InitActiveObject()
        {
            // 必须先运行幻灯片，下面才能获得PowerPoint应用程序，否则会出现异常
            // 获得正在运行的PowerPoint应用程序
            try
            {
                _pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                if (_pptApplication != null)
                {
                    _isGetPptObjectCompleted = true;
                    Init();
                    _pptApplication.SlideShowBegin += PptApplication_SlideShowBegin;
                    _pptApplication.SlideShowEnd += PptApplication_SlideShowEnd;
                    _pptApplication.SlideShowNextSlide += PptApplication_SlideShowNextSlide;

                }
            }
            catch (Exception ce)
            {
                new Exception("请先启动遥控的幻灯片", ce);
            }
        }

        protected virtual void OnPageChanged(int pageNo)
        {
            try
            {
                PageChanged?.Invoke(pageNo);
            }
            catch (Exception)
            {
                //ignore
            }
        }

        protected virtual void OnSlideShowEnd()
        {
            try
            {
                SlideShowEnd?.Invoke();
            }
            catch (Exception)
            {
                //ignore
            }
        }

        public void Run()
        {
            if (_pptApplication != null)
            {
                // 获得当前选中的幻灯片
                try
                {
                    // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                    // 然而在阅读模式下，这种方式会出现异常
                    slide = slides[_pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch (Exception)
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = _pptApplication.SlideShowWindows[1].View.Slide;
                }
                CheckRun();
            }
        }

        private void Init()
        {
            //获得演示文稿对象
            _presentation = _pptApplication.ActivePresentation;
            // 获得幻灯片对象集合
            slides = _presentation.Slides;
            // 获得幻灯片的数量
            PageCount = slides.Count;

        }

        private void CheckRun()
        {
            if (_presentation != null)
            {
                var objSSS = _presentation.SlideShowSettings;
                //如过你不想循环放映就把TRUE改成FALSE. 
                objSSS.LoopUntilStopped = MsoTriState.msoTrue;
                //objSSS.StartingSlide = 16;
                //objSSS.EndingSlide = PageCount;
                //objSSS.Run(); //Wait for the slide show to end. 
            }
        }

        public void First()
        {
            try
            {
                // 在普通视图中调用Select方法来选中第一张幻灯片
                slides[1].Select();
                slide = slides[1];

                #region 此段代码为了兼容WPS
                _pptApplication?.SlideShowWindows[1].View.First();
                slide = _pptApplication?.SlideShowWindows[1].View.Slide;
                #endregion
            }
            catch
            {
                try
                {
                    // 在阅读模式下使用下面的方式来切换到第一张幻灯片
                    _pptApplication?.SlideShowWindows[1].View.First();
                    slide = _pptApplication?.SlideShowWindows[1].View.Slide;
                }
                catch (Exception)
                {
                    //ignore
                }
            }
        }

        public void Previous()
        {
            try
            {
                _presentation?.SlideShowWindow.View.Previous();
            }
            catch (Exception)
            {
                //ignore
            }
        }

        public void Next()
        {
            try
            {
                //CheckRun();//由用户手动切换到放映态
                _presentation?.SlideShowWindow.View.Next();
            }
            catch (Exception)
            {
                //ignore
            }
        }

        public void Last()
        {
            try
            {
                _presentation?.SlideShowWindow.View.Last();
            }
            catch (Exception)
            {
                //ignore
            }
        }

        public void GotoPage(int pageNumber)
        {
            if (pageNumber <= 0 || pageNumber > PageCount) throw new ArgumentOutOfRangeException(nameof(pageNumber));
            try
            {
                _presentation?.SlideShowWindow.View.GotoSlide(pageNumber);
            }
            catch (Exception)
            {
                //ignore
            }
        }

        public void Exit()
        {
            try
            {
                _presentation?.SlideShowWindow.View.Exit();
            }
            catch (COMException)
            {
                //ignore
            }
        }

        #region 事件处理方法
        private void PptApplication_SlideShowEnd(PPt.Presentation Pres)
        {
            Dispose();
        }

        private void PptApplication_SlideShowBegin(PPt.SlideShowWindow Wn)
        {

        }
        private void PptApplication_SlideShowNextSlide(PPt.SlideShowWindow Wn)
        {
            int p = Wn.View.CurrentShowPosition;
            try
            {
                OnPageChanged(p);
            }
            catch (Exception)
            {
                //ignore
            }
        }
        #endregion

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            if (_pptApplication != null)
            {
                _pptApplication.SlideShowBegin -= PptApplication_SlideShowBegin;
                _pptApplication.SlideShowEnd -= PptApplication_SlideShowEnd;
                _pptApplication.SlideShowNextSlide -= PptApplication_SlideShowNextSlide;
                //_pptApplication.Quit();此处不能用Quit,否则用户主动退出放映态时得不到事件通知
                Marshal.ReleaseComObject(_presentation);
                Marshal.ReleaseComObject(_pptApplication);
                if (slide != null)
                    Marshal.ReleaseComObject(slide);
                Marshal.ReleaseComObject(slides);
                _pptApplication = null;
                _presentation = null;
                slide = null;
                slides = null;
            }

            try
            {
                OnSlideShowEnd();

            }
            catch (Exception)
            {
                //ignore
            }
        }
    }
}
