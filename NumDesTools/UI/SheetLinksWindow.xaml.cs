using GraphX.Common.Enums;
using GraphX.Controls;
using GraphX.Logic.Models;
using QuickGraph;
using System.Windows.Input;
using MessageBox = System.Windows.MessageBox;

namespace NumDesTools.UI
{
    /// <summary>
    /// SheetLinksWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SheetLinksWindow
    {
        public SheetLinksWindow()
        {
            InitializeComponent();


            // 创建图形
            var graph = new BidirectionalGraph<SelfGraphXVertex, SelfGraphXEdge>();

            // 添加顶点
            var v1 = new SelfGraphXVertex { Name = "Vertex 1" };
            var v2 = new SelfGraphXVertex { Name = "Vertex 2" };
            graph.AddVertex(v1);
            graph.AddVertex(v2);

            // 添加边
            var e1 = new SelfGraphXEdge(v1, v2);
            graph.AddEdge(e1);

            // 创建图形控件
            var graphArea = new GraphArea<SelfGraphXVertex, SelfGraphXEdge, BidirectionalGraph<SelfGraphXVertex, SelfGraphXEdge>>();
            // 创建并设置LogicCore
            var logicCore = new GXLogicCore<SelfGraphXVertex, SelfGraphXEdge, BidirectionalGraph<SelfGraphXVertex, SelfGraphXEdge>>();
            logicCore.Graph = graph;
            graphArea.LogicCore = logicCore;

            // 布局和渲染图形
            graphArea.LogicCore.DefaultLayoutAlgorithm = LayoutAlgorithmTypeEnum.KK;
            graphArea.GenerateGraph();

            // 为顶点和边添加交互事件，并设置顶点的标签
            foreach (var vc in graphArea.VertexList)
            {
                vc.Value.MouseLeftButtonDown += VertexControl_MouseLeftButtonDown;
                vc.Value.MouseRightButtonDown += VertexControl_MouseRightButtonDown;
            }

            foreach (var ec in graphArea.EdgesList)
            {
                ec.Value.MouseLeftButtonDown += EdgeControl_MouseLeftButtonDown;
                ec.Value.MouseRightButtonDown += EdgeControl_MouseRightButtonDown;
            }


            // 创建一个ZoomControl并将GraphArea添加到其中
            ZoomControl zoomControl = new ZoomControl();
            zoomControl.Content = graphArea;

            BaseGrid.Children.Add(zoomControl);
        }
        private void VertexControl_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            VertexControl vc = (VertexControl)sender;
            SelfGraphXVertex vertex = (SelfGraphXVertex)vc.Vertex;
            MessageBox.Show($"点击了顶点：{vertex.Name}");
        }

        private void VertexControl_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            VertexControl vc = (VertexControl)sender;
            SelfGraphXVertex vertex = (SelfGraphXVertex)vc.Vertex;
            MessageBox.Show($"右击了顶点：{vertex.Name}");
        }

        private void EdgeControl_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            EdgeControl ec = (EdgeControl)sender;
            SelfGraphXEdge edge = (SelfGraphXEdge)ec.Edge;
            MessageBox.Show($"点击了边：{edge.Source.Name} -> {edge.Target.Name}");
        }

        private void EdgeControl_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            EdgeControl ec = (EdgeControl)sender;
            SelfGraphXEdge edge = (SelfGraphXEdge)ec.Edge;
            MessageBox.Show($"右击了边：{edge.Source.Name} -> {edge.Target.Name}");
        }
    }

}
