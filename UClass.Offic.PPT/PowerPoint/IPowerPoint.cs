using System;

namespace UClass.Office.PowerPoint
{
    public interface IPowerPoint : IDisposable
    {
        /// <summary>
        /// ppt退出播放模式时触发
        /// </summary>
        event Action SlideShowEnd;
        /// <summary>
        /// 翻页事件,参数为当前页索引(从1开始)
        /// </summary>
        event Action<int> PageChanged;

        /// <summary>
        /// 初始化PPT对象是否完成.
        /// </summary>
        bool IsGetPptObjectCompleted { get; }

        bool IsDisposed { get; }

        /// <summary>
        /// PPT总页数
        /// </summary>
        int PageCount { get; }
        /// <summary>
        /// 当前页
        /// </summary>
        int CurrentPage { get; }
        /// <summary>
        /// 首先要初始化PPT对象
        /// </summary>
        void InitActiveObject();
        /// <summary>
        /// 启动PPT遥控
        /// </summary>
        void Run();
        /// <summary>
        /// 显示第一页
        /// </summary>
        void First();
        /// <summary>
        /// 上一页
        /// </summary>
        void Previous();
        /// <summary>
        /// 下一页
        /// </summary>
        void Next();
        /// <summary>
        /// 尾页
        /// </summary>
        void Last();
        /// <summary>
        /// 跳转到指定页
        /// </summary>
        /// <param name="pageNumber"></param>
        void GotoPage(int pageNumber);
        /// <summary>
        /// 退出播放模式
        /// </summary>
        void Exit();
    }
}
