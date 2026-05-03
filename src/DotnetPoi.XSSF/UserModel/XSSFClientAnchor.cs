namespace DotnetPoi.XSSF.UserModel;

public enum AnchorType
{
    MOVE_AND_RESIZE,
    DONT_MOVE_AND_RESIZE,
    MOVE_DONT_RESIZE
}

public sealed class XSSFClientAnchor
{
    private AnchorType _anchorType = AnchorType.MOVE_AND_RESIZE;

    public XSSFClientAnchor()
        : this(0, 0, 0, 0, 0, 0, 0, 0)
    {
    }

    public XSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
    {
        setDx1(dx1);
        setDy1(dy1);
        setDx2(dx2);
        setDy2(dy2);
        setCol1(col1);
        setRow1(row1);
        setCol2(col2);
        setRow2(row2);
    }

    public int getDx1() => Dx1;

    public void setDx1(int dx1) => Dx1 = dx1;

    public int getDy1() => Dy1;

    public void setDy1(int dy1) => Dy1 = dy1;

    public int getDx2() => Dx2;

    public void setDx2(int dx2) => Dx2 = dx2;

    public int getDy2() => Dy2;

    public void setDy2(int dy2) => Dy2 = dy2;

    public short getCol1() => (short)Col1;

    public void setCol1(int col1) => Col1 = col1;

    public int getRow1() => Row1;

    public void setRow1(int row1) => Row1 = row1;

    public short getCol2() => (short)Col2;

    public void setCol2(int col2) => Col2 = col2;

    public int getRow2() => Row2;

    public void setRow2(int row2) => Row2 = row2;

    public void setAnchorType(AnchorType anchorType) => _anchorType = anchorType;

    public AnchorType getAnchorType() => _anchorType;

    internal int Dx1 { get; private set; }

    internal int Dy1 { get; private set; }

    internal int Dx2 { get; private set; }

    internal int Dy2 { get; private set; }

    internal int Col1 { get; private set; }

    internal int Row1 { get; private set; }

    internal int Col2 { get; private set; }

    internal int Row2 { get; private set; }
}
