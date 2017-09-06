/**
 * Created by Cser_W on 2017/9/6.
 */
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.com.ComThread;


public class WordObj
{

    public WordObj()
    {
    }

    private static WordObj instance;

    private Dispatch doc = null;

    private Dispatch activeWindow = null;

    private Dispatch docSelection = null;

    private Dispatch wrdDocs = null;

    private String fileName;

    private ActiveXComponent wrdCom;

    /**
     * 获取Word操作静态实例对象
     *
     * @return 报表汇总业务操作
     */
    public final static synchronized WordObj getInstance()
    {
        if (instance == null)
            instance = new WordObj();
        return instance;

    }

    /**
     * 初始化Word对象
     *
     * @return 是否初始化成功
     */
    public boolean initWordObj()
    {
        boolean retFlag = false;
        ComThread.InitSTA();// 初始化com的线程，非常重要！！使用结束后要调用 realease方法
        wrdCom = new ActiveXComponent("Word.Application");
        try
        {
            // 返回wrdCom.Documents的Dispatch
            wrdDocs = wrdCom.getProperty("Documents").toDispatch();
            wrdCom.setProperty("Visible", new Variant(true));

            retFlag = true;
        }
        catch (Exception e)
        {
            retFlag = false;
            e.printStackTrace();
        }
        return retFlag;
    }

    /**
     * 创建一个新的word文档
     *
     */

    public void createNewDocument()
    {
        doc = Dispatch.call(wrdDocs, "Add").toDispatch();
        docSelection = Dispatch.get(wrdCom, "Selection").toDispatch();

    }

    /**
     * 取得活动窗体对象
     *
     */
    public void getActiveWindow()
    {
        // 取得活动窗体对象
        activeWindow = wrdCom.getProperty("ActiveWindow").toDispatch();

    }

    /**
     * 打开一个已存在的文档
     *
     * @param docPath
     */

    public void openDocument(String docPath)
    {
        if (this.doc != null)
        {
            this.closeDocument();
        }
        doc = Dispatch.call(wrdDocs, "Open", docPath).toDispatch();
        docSelection = Dispatch.get(wrdCom, "Selection").toDispatch();
    }

    /**
     * 关闭当前word文档
     *
     */
    public void closeDocument()
    {
        if (doc != null)
        {
            Dispatch.call(doc, "Save");
            Dispatch.call(doc, "Close", new Variant(0));
            doc = null;
        }
    }

    /**
     * 文档设置水印
     *
     * @param waterMarkStr 水印字符串
     */
    public void setWaterMark(String waterMarkStr)
    {
        // 取得活动窗格对象
        Dispatch activePan = Dispatch.get(activeWindow, "ActivePane")
                .toDispatch();
        // 取得视窗对象
        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
        //输入页眉内容
        Dispatch.put(view, "SeekView", new Variant(9));
        Dispatch headfooter = Dispatch.get(docSelection, "HeaderFooter")
                .toDispatch();
        //取得图形对象
        Dispatch shapes = Dispatch.get(headfooter, "Shapes").toDispatch();
        //给文档全部加上水印
        Dispatch selection = Dispatch.call(shapes, "AddTextEffect",
                new Variant(9), waterMarkStr, "宋体", new Variant(1),
                new Variant(false), new Variant(false), new Variant(0),
                new Variant(0)).toDispatch();
        Dispatch.call(selection, "Select");
        //设置水印参数
        Dispatch shapeRange = Dispatch.get(docSelection, "ShapeRange")
                .toDispatch();
        Dispatch.put(shapeRange, "Name", "PowerPlusWaterMarkObject1");
        Dispatch textEffect = Dispatch.get(shapeRange, "TextEffect").toDispatch();
        Dispatch.put(textEffect, "NormalizedHeight", new Boolean(false));
        Dispatch line = Dispatch.get(shapeRange, "Line").toDispatch();
        Dispatch.put(line, "Visible", new Boolean(false));
        Dispatch fill = Dispatch.get(shapeRange, "Fill").toDispatch();
        Dispatch.put(fill, "Visible", new Boolean(true));
        //设置水印透明度
        Dispatch.put(fill, "Transparency", new Variant(0.5));
        Dispatch foreColor = Dispatch.get(fill, "ForeColor").toDispatch();
        //设置水印颜色
        Dispatch.put(foreColor, "RGB", new Variant(16711680));
        Dispatch.call(fill, "Solid");
        //设置水印旋转
        Dispatch.put(shapeRange, "Rotation", new Variant(315));
        Dispatch.put(shapeRange, "LockAspectRatio", new Boolean(true));
        Dispatch.put(shapeRange, "Height", new Variant(117.0709));
        Dispatch.put(shapeRange, "Width", new Variant(468.2835));
        Dispatch.put(shapeRange, "Left", new Variant(-999995));
        Dispatch.put(shapeRange, "Top", new Variant(-999995));
        Dispatch wrapFormat = Dispatch.get(shapeRange, "WrapFormat").toDispatch();
        //是否允许交叠
        Dispatch.put(wrapFormat, "AllowOverlap", new Variant(true));
        Dispatch.put(wrapFormat, "Side", new Variant(3));
        Dispatch.put(wrapFormat, "Type", new Variant(3));
        Dispatch.put(shapeRange, "RelativeHorizontalPosition", new Variant(0));
        Dispatch.put(shapeRange, "RelativeVerticalPosition", new Variant(0));
        Dispatch.put(view, "SeekView", new Variant(0));
    }



    /**
     * 关闭Word资源
     *
     *
     */
    public void closeWordObj()
    {

        // 关闭word文件
        wrdCom.invoke("Quit", new Variant[] {});
        // 释放com线程。根据jacob的帮助文档，com的线程回收不由java的垃圾回收器处理
        ComThread.Release();
    }

    /**
     * 得到文件名
     *
     * @return .
     */
    public String getFileName()
    {
        return fileName;
    }

    /**
     * 设置文件名
     *
     * @param fileName .
     */
    public void setFileName(String fileName)
    {
        this.fileName = fileName;
    }

    /**
     * 测试功能
     *
     */
    public static void main(String[] argv)
    {
        WordObj d = WordObj.getInstance();
        try
        {
            if (d.initWordObj())
            {
              //  d.createNewDocument();
                String path = "C:\\Users\\Cser_W\\Desktop\\王兴银研二第一学期工作计划 (1).docx";
                d.openDocument(path);
                d.getActiveWindow();
                d.setWaterMark("HEllo,this is ok!");
                 d.closeWordObj();
            }
            else
                System.out.println("初始化Word读写对象失败！");
        }
        catch (Exception e)
        {
            d.closeWordObj();
        }
    }
}
