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
     * ��ȡWord������̬ʵ������
     *
     * @return �������ҵ�����
     */
    public final static synchronized WordObj getInstance()
    {
        if (instance == null)
            instance = new WordObj();
        return instance;

    }

    /**
     * ��ʼ��Word����
     *
     * @return �Ƿ��ʼ���ɹ�
     */
    public boolean initWordObj()
    {
        boolean retFlag = false;
        ComThread.InitSTA();// ��ʼ��com���̣߳��ǳ���Ҫ����ʹ�ý�����Ҫ���� realease����
        wrdCom = new ActiveXComponent("Word.Application");
        try
        {
            // ����wrdCom.Documents��Dispatch
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
     * ����һ���µ�word�ĵ�
     *
     */

    public void createNewDocument()
    {
        doc = Dispatch.call(wrdDocs, "Add").toDispatch();
        docSelection = Dispatch.get(wrdCom, "Selection").toDispatch();

    }

    /**
     * ȡ�û�������
     *
     */
    public void getActiveWindow()
    {
        // ȡ�û�������
        activeWindow = wrdCom.getProperty("ActiveWindow").toDispatch();

    }

    /**
     * ��һ���Ѵ��ڵ��ĵ�
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
     * �رյ�ǰword�ĵ�
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
     * �ĵ�����ˮӡ
     *
     * @param waterMarkStr ˮӡ�ַ���
     */
    public void setWaterMark(String waterMarkStr)
    {
        // ȡ�û�������
        Dispatch activePan = Dispatch.get(activeWindow, "ActivePane")
                .toDispatch();
        // ȡ���Ӵ�����
        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
        //����ҳü����
        Dispatch.put(view, "SeekView", new Variant(9));
        Dispatch headfooter = Dispatch.get(docSelection, "HeaderFooter")
                .toDispatch();
        //ȡ��ͼ�ζ���
        Dispatch shapes = Dispatch.get(headfooter, "Shapes").toDispatch();
        //���ĵ�ȫ������ˮӡ
        Dispatch selection = Dispatch.call(shapes, "AddTextEffect",
                new Variant(9), waterMarkStr, "����", new Variant(1),
                new Variant(false), new Variant(false), new Variant(0),
                new Variant(0)).toDispatch();
        Dispatch.call(selection, "Select");
        //����ˮӡ����
        Dispatch shapeRange = Dispatch.get(docSelection, "ShapeRange")
                .toDispatch();
        Dispatch.put(shapeRange, "Name", "PowerPlusWaterMarkObject1");
        Dispatch textEffect = Dispatch.get(shapeRange, "TextEffect").toDispatch();
        Dispatch.put(textEffect, "NormalizedHeight", new Boolean(false));
        Dispatch line = Dispatch.get(shapeRange, "Line").toDispatch();
        Dispatch.put(line, "Visible", new Boolean(false));
        Dispatch fill = Dispatch.get(shapeRange, "Fill").toDispatch();
        Dispatch.put(fill, "Visible", new Boolean(true));
        //����ˮӡ͸����
        Dispatch.put(fill, "Transparency", new Variant(0.5));
        Dispatch foreColor = Dispatch.get(fill, "ForeColor").toDispatch();
        //����ˮӡ��ɫ
        Dispatch.put(foreColor, "RGB", new Variant(16711680));
        Dispatch.call(fill, "Solid");
        //����ˮӡ��ת
        Dispatch.put(shapeRange, "Rotation", new Variant(315));
        Dispatch.put(shapeRange, "LockAspectRatio", new Boolean(true));
        Dispatch.put(shapeRange, "Height", new Variant(117.0709));
        Dispatch.put(shapeRange, "Width", new Variant(468.2835));
        Dispatch.put(shapeRange, "Left", new Variant(-999995));
        Dispatch.put(shapeRange, "Top", new Variant(-999995));
        Dispatch wrapFormat = Dispatch.get(shapeRange, "WrapFormat").toDispatch();
        //�Ƿ�������
        Dispatch.put(wrapFormat, "AllowOverlap", new Variant(true));
        Dispatch.put(wrapFormat, "Side", new Variant(3));
        Dispatch.put(wrapFormat, "Type", new Variant(3));
        Dispatch.put(shapeRange, "RelativeHorizontalPosition", new Variant(0));
        Dispatch.put(shapeRange, "RelativeVerticalPosition", new Variant(0));
        Dispatch.put(view, "SeekView", new Variant(0));
    }



    /**
     * �ر�Word��Դ
     *
     *
     */
    public void closeWordObj()
    {

        // �ر�word�ļ�
        wrdCom.invoke("Quit", new Variant[] {});
        // �ͷ�com�̡߳�����jacob�İ����ĵ���com���̻߳��ղ���java����������������
        ComThread.Release();
    }

    /**
     * �õ��ļ���
     *
     * @return .
     */
    public String getFileName()
    {
        return fileName;
    }

    /**
     * �����ļ���
     *
     * @param fileName .
     */
    public void setFileName(String fileName)
    {
        this.fileName = fileName;
    }

    /**
     * ���Թ���
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
                String path = "C:\\Users\\Cser_W\\Desktop\\�������ж���һѧ�ڹ����ƻ� (1).docx";
                d.openDocument(path);
                d.getActiveWindow();
                d.setWaterMark("HEllo,this is ok!");
                 d.closeWordObj();
            }
            else
                System.out.println("��ʼ��Word��д����ʧ�ܣ�");
        }
        catch (Exception e)
        {
            d.closeWordObj();
        }
    }
}
