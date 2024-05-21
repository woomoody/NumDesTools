using GraphX.Common.Enums;
using GraphX.Controls;
using GraphX.Logic.Algorithms.LayoutAlgorithms;
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


            //// 创建图形
            //var graph = new CompoundGraph<SelfGraphXVertex, SelfGraphXEdge>();

            //// 添加顶点
            //var v1 = new SelfGraphXVertex { Name = "Vertex 1" };
            //var v2 = new SelfGraphXVertex { Name = "Vertex 2" };
            //var v3 = new SelfGraphXVertex { Name = "Vertex 3" };
            //graph.AddVertex(v1);
            //graph.AddVertex(v2);
            //graph.AddChildVertex(v1, v3);  // 添加一个子顶点

            //// 添加边
            //var e1 = new SelfGraphXEdge(v1, v2);
            //graph.AddEdge(e1);


            // 创建一个 CompoundGraph 实例
            var graph = new CompoundGraph<SelfGraphXVertex, SelfGraphXEdge>();

            // 创建父节点 A 和 B
            var parentA = new SelfGraphXVertex { Name = "Parent A" };
            var parentB = new SelfGraphXVertex { Name = "Parent B" };

            // 创建 A 的子节点
            var childA1 = new SelfGraphXVertex { Name = "Child A1" };
            var childA2 = new SelfGraphXVertex { Name = "Child A2" };
            var childA3 = new SelfGraphXVertex { Name = "Child A3" };

            // 创建 B 的子节点
            var childB1 = new SelfGraphXVertex { Name = "Child B1" };
            var childB2 = new SelfGraphXVertex { Name = "Child B2" };

            // 添加父节点和子节点
            graph.AddVertex(parentA);
            graph.AddChildVertex(parentA, childA1);
            graph.AddChildVertex(parentA, childA2);
            graph.AddChildVertex(parentA, childA3);

            graph.AddVertex(parentB);
            graph.AddChildVertex(parentB, childB1);
            graph.AddChildVertex(parentB, childB2);

            // 创建一条从 A 的子节点 2 到 B 的子节点 1 的边
            var edge = new SelfGraphXEdge(childA2, childB1);
            graph.AddEdge(edge);


            // 创建图形控件
            var graphArea = new GraphArea<SelfGraphXVertex, SelfGraphXEdge, CompoundGraph<SelfGraphXVertex, SelfGraphXEdge>>();
            // 创建并设置LogicCore
            var logicCore = new GXLogicCore<SelfGraphXVertex, SelfGraphXEdge, CompoundGraph<SelfGraphXVertex, SelfGraphXEdge>>();
            logicCore.Graph = graph;
            graphArea.LogicCore = logicCore;

            // 布局和渲染图形
            graphArea.LogicCore.DefaultLayoutAlgorithm = LayoutAlgorithmTypeEnum.ISOM;
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
