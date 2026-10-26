{ ============================================================================
  ufrmsnap - the window as a PNG, for --snapshot=<file.png> test runs.
  GetFormImage paints only the windowed controls, so labels, shapes and
  group box captions are drawn on top of it here.
  ============================================================================ }
unit ufrmsnap;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Math, Types, Forms, Controls, Graphics, StdCtrls, ExtCtrls, ComCtrls, Grids;

procedure SaveFormSnapshot(AForm: TCustomForm; const FileName: string);

implementation

procedure SaveFormSnapshot(AForm: TCustomForm; const FileName: string);
var
  Bmp: TBitmap;
  Png: TPortableNetworkGraphic;

  procedure DrawGrid(G: TStringGrid; P: TPoint);
  var
    R, Col, X, Y, H: Integer;
    Cell: TRect;
  begin
    Bmp.Canvas.Font := G.Font;
    Bmp.Canvas.Pen.Color := $E0E0E0;
    H := G.DefaultRowHeight;
    Y := P.Y;
    for R := 0 to G.RowCount - 1 do
    begin
      if (R >= G.FixedRows) and (R < G.TopRow) then Continue;
      if Y + H > P.Y + G.ClientHeight then Break;
      X := P.X;
      for Col := 0 to G.ColCount - 1 do
      begin
        Cell := Rect(X, Y, Min(X + G.ColWidths[Col], P.X + G.ClientWidth), Y + H);
        if R < G.FixedRows then
          Bmp.Canvas.Brush.Color := clBtnFace
        else
          Bmp.Canvas.Brush.Color := clWindow;
        Bmp.Canvas.Rectangle(Cell.Left, Cell.Top, Cell.Right + 1, Cell.Bottom + 1);
        Bmp.Canvas.TextRect(Rect(Cell.Left + 3, Cell.Top, Cell.Right - 2, Cell.Bottom),
          Cell.Left + 3, Cell.Top + (H - Bmp.Canvas.TextHeight('Xg')) div 2, G.Cells[Col, R]);
        Inc(X, G.ColWidths[Col]);
        if X >= P.X + G.ClientWidth then Break;
      end;
      Inc(Y, H);
    end;
  end;

  procedure DrawGraphics(Parent: TWinControl);
  var
    I: Integer;
    C: TControl;
    P: TPoint;
  begin
    for I := 0 to Parent.ControlCount - 1 do
    begin
      C := Parent.Controls[I];
      if not C.Visible then Continue;
      P := C.ClientToParent(Point(0, 0), AForm);
      if C is TLabel then
      begin
        Bmp.Canvas.Font := TLabel(C).Font;
        Bmp.Canvas.Brush.Style := bsClear;
        Bmp.Canvas.TextOut(P.X, P.Y, TLabel(C).Caption);
      end
      else if C is TShape then
      begin
        Bmp.Canvas.Pen.Color := clGray;
        Bmp.Canvas.Brush.Style := bsSolid;
        Bmp.Canvas.Brush.Color := TShape(C).Brush.Color;
        Bmp.Canvas.Ellipse(P.X, P.Y, P.X + C.Width, P.Y + C.Height);
      end
      else if C is TWinControl then
      begin
        { a grid on a tab sheet is not in the form image (the tab control
          does not print its children): draw its cells }
        if (C is TStringGrid) and (C.Parent is TTabSheet) and TWinControl(C).Showing then
          DrawGrid(TStringGrid(C), P);
        if C is TGroupBox then
        begin
          Bmp.Canvas.Font := TGroupBox(C).Font;
          Bmp.Canvas.Brush.Style := bsClear;
          Bmp.Canvas.Pen.Color := clSilver;
          Bmp.Canvas.Rectangle(P.X + 2, P.Y + 7, P.X + C.Width - 2, P.Y + C.Height - 2);
          Bmp.Canvas.Brush.Style := bsSolid;
          Bmp.Canvas.Brush.Color := clBtnFace;
          Bmp.Canvas.TextOut(P.X + 10, P.Y, ' ' + TGroupBox(C).Caption + ' ');
        end;
        DrawGraphics(TWinControl(C));
      end;
    end;
  end;

begin
  Bmp := AForm.GetFormImage;
  Png := TPortableNetworkGraphic.Create;
  try
    if Bmp.PixelFormat <> pf24bit then Bmp.PixelFormat := pf24bit;
    DrawGraphics(AForm);
    Png.Assign(Bmp);
    Png.SaveToFile(FileName);
  finally
    Png.Free;
    Bmp.Free;
  end;
end;

end.
