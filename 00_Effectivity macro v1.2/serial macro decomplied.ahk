; <COMPILER: v1.1.23.01>
#NoEnv
SendMode Input
SetWorkingDir   %A_ScriptDir%
#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#KeyHistory 0
Process, Priority, , H
SetBatchLines, -1
SetKeyDelay, -1, -1, Play
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, 0
SetControlDelay, 0
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
#SingleInstance Force
++SetTitleMatchMode, 2
DetectHiddenWindows On
DetectHiddenText on
UpdateLayeredWindow(hwnd, hdc, x="", y="", w="", h="", Alpha=255)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if ((x != "") && (y != ""))
VarSetCapacity(pt, 8), NumPut(x, pt, 0, "UInt"), NumPut(y, pt, 4, "UInt")
if (w = "") ||(h = "")
WinGetPos,,, w, h, ahk_id %hwnd%
return DllCall("UpdateLayeredWindow"
, Ptr, hwnd
, Ptr, 0
, Ptr, ((x = "") && (y = "")) ? 0 : &pt
, "int64*", w|h<<32
, Ptr, hdc
, "int64*", 0
, "uint", 0
, "UInt*", Alpha<<16|1<<24
, "uint", 2)
}
BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, Raster="")
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdi32\BitBlt"
, Ptr, dDC
, "int", dx
, "int", dy
, "int", dw
, "int", dh
, Ptr, sDC
, "int", sx
, "int", sy
, "uint", Raster ? Raster : 0x00CC0020)
}
StretchBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, sw, sh, Raster="")
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdi32\StretchBlt"
, Ptr, ddc
, "int", dx
, "int", dy
, "int", dw
, "int", dh
, Ptr, sdc
, "int", sx
, "int", sy
, "int", sw
, "int", sh
, "uint", Raster ? Raster : 0x00CC0020)
}
SetStretchBltMode(hdc, iStretchMode=4)
{
return DllCall("gdi32\SetStretchBltMode"
, A_PtrSize ? "UPtr" : "UInt", hdc
, "int", iStretchMode)
}
SetImage(hwnd, hBitmap)
{
SendMessage, 0x172, 0x0, hBitmap,, ahk_id %hwnd%
E := ErrorLevel
DeleteObject(E)
return E
}
SetSysColorToControl(hwnd, SysColor=15)
{
WinGetPos,,, w, h, ahk_id %hwnd%
bc := DllCall("GetSysColor", "Int", SysColor, "UInt")
pBrushClear := Gdip_BrushCreateSolid(0xff000000 | (bc >> 16 | bc & 0xff00 | (bc & 0xff) << 16))
pBitmap := Gdip_CreateBitmap(w, h), G := Gdip_GraphicsFromImage(pBitmap)
Gdip_FillRectangle(G, pBrushClear, 0, 0, w, h)
hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
SetImage(hwnd, hBitmap)
Gdip_DeleteBrush(pBrushClear)
Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmap), DeleteObject(hBitmap)
return 0
}
Gdip_BitmapFromScreen(Screen=0, Raster="")
{
if (Screen = 0)
{
Sysget, x, 76
Sysget, y, 77
Sysget, w, 78
Sysget, h, 79
}
else if (SubStr(Screen, 1, 5) = "hwnd:")
{
Screen := SubStr(Screen, 6)
if !WinExist( "ahk_id " Screen)
return -2
WinGetPos,,, w, h, ahk_id %Screen%
x := y := 0
hhdc := GetDCEx(Screen, 3)
}
else if (Screen&1 != "")
{
Sysget, M, Monitor, %Screen%
x := MLeft, y := MTop, w := MRight-MLeft, h := MBottom-MTop
}
else
{
StringSplit, S, Screen, |
x := S1, y := S2, w := S3, h := S4
}
if (x = "") || (y = "") || (w = "") || (h = "")
return -1
chdc := CreateCompatibleDC(), hbm := CreateDIBSection(w, h, chdc), obm := SelectObject(chdc, hbm), hhdc := hhdc ? hhdc : GetDC()
BitBlt(chdc, 0, 0, w, h, hhdc, x, y, Raster)
ReleaseDC(hhdc)
pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
SelectObject(chdc, obm), DeleteObject(hbm), DeleteDC(hhdc), DeleteDC(chdc)
return pBitmap
}
Gdip_BitmapFromHWND(hwnd)
{
WinGetPos,,, Width, Height, ahk_id %hwnd%
hbm := CreateDIBSection(Width, Height), hdc := CreateCompatibleDC(), obm := SelectObject(hdc, hbm)
PrintWindow(hwnd, hdc)
pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
SelectObject(hdc, obm), DeleteObject(hbm), DeleteDC(hdc)
return pBitmap
}
CreateRectF(ByRef RectF, x, y, w, h)
{
VarSetCapacity(RectF, 16)
NumPut(x, RectF, 0, "float"), NumPut(y, RectF, 4, "float"), NumPut(w, RectF, 8, "float"), NumPut(h, RectF, 12, "float")
}
CreateRect(ByRef Rect, x, y, w, h)
{
VarSetCapacity(Rect, 16)
NumPut(x, Rect, 0, "uint"), NumPut(y, Rect, 4, "uint"), NumPut(w, Rect, 8, "uint"), NumPut(h, Rect, 12, "uint")
}
CreateSizeF(ByRef SizeF, w, h)
{
VarSetCapacity(SizeF, 8)
NumPut(w, SizeF, 0, "float"), NumPut(h, SizeF, 4, "float")
}
CreatePointF(ByRef PointF, x, y)
{
VarSetCapacity(PointF, 8)
NumPut(x, PointF, 0, "float"), NumPut(y, PointF, 4, "float")
}
CreateDIBSection(w, h, hdc="", bpp=32, ByRef ppvBits=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
hdc2 := hdc ? hdc : GetDC()
VarSetCapacity(bi, 40, 0)
NumPut(w, bi, 4, "uint")
, NumPut(h, bi, 8, "uint")
, NumPut(40, bi, 0, "uint")
, NumPut(1, bi, 12, "ushort")
, NumPut(0, bi, 16, "uInt")
, NumPut(bpp, bi, 14, "ushort")
hbm := DllCall("CreateDIBSection"
, Ptr, hdc2
, Ptr, &bi
, "uint", 0
, A_PtrSize ? "UPtr*" : "uint*", ppvBits
, Ptr, 0
, "uint", 0, Ptr)
if !hdc
ReleaseDC(hdc2)
return hbm
}
PrintWindow(hwnd, hdc, Flags=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("PrintWindow", Ptr, hwnd, Ptr, hdc, "uint", Flags)
}
DestroyIcon(hIcon)
{
return DllCall("DestroyIcon", A_PtrSize ? "UPtr" : "UInt", hIcon)
}
PaintDesktop(hdc)
{
return DllCall("PaintDesktop", A_PtrSize ? "UPtr" : "UInt", hdc)
}
CreateCompatibleBitmap(hdc, w, h)
{
return DllCall("gdi32\CreateCompatibleBitmap", A_PtrSize ? "UPtr" : "UInt", hdc, "int", w, "int", h)
}
CreateCompatibleDC(hdc=0)
{
return DllCall("CreateCompatibleDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}
SelectObject(hdc, hgdiobj)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("SelectObject", Ptr, hdc, Ptr, hgdiobj)
}
DeleteObject(hObject)
{
return DllCall("DeleteObject", A_PtrSize ? "UPtr" : "UInt", hObject)
}
GetDC(hwnd=0)
{
return DllCall("GetDC", A_PtrSize ? "UPtr" : "UInt", hwnd)
}
GetDCEx(hwnd, flags=0, hrgnClip=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("GetDCEx", Ptr, hwnd, Ptr, hrgnClip, "int", flags)
}
ReleaseDC(hdc, hwnd=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("ReleaseDC", Ptr, hwnd, Ptr, hdc)
}
DeleteDC(hdc)
{
return DllCall("DeleteDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}
Gdip_LibraryVersion()
{
return 1.45
}
Gdip_LibrarySubVersion()
{
return 1.47
}
Gdip_BitmapFromBRA(ByRef BRAFromMemIn, File, Alternate=0)
{
Static FName = "ObjRelease"
if !BRAFromMemIn
return -1
Loop, Parse, BRAFromMemIn, `n
{
if (A_Index = 1)
{
StringSplit, Header, A_LoopField, |
if (Header0 != 4 || Header2 != "BRA!")
return -2
}
else if (A_Index = 2)
{
StringSplit, Info, A_LoopField, |
if (Info0 != 3)
return -3
}
else
break
}
if !Alternate
StringReplace, File, File, \, \\, All
RegExMatch(BRAFromMemIn, "mi`n)^" (Alternate ? File "\|.+?\|(\d+)\|(\d+)" : "\d+\|" File "\|(\d+)\|(\d+)") "$", FileInfo)
if !FileInfo
return -4
hData := DllCall("GlobalAlloc", "uint", 2, Ptr, FileInfo2, Ptr)
pData := DllCall("GlobalLock", Ptr, hData, Ptr)
DllCall("RtlMoveMemory", Ptr, pData, Ptr, &BRAFromMemIn+Info2+FileInfo1, Ptr, FileInfo2)
DllCall("GlobalUnlock", Ptr, hData)
DllCall("ole32\CreateStreamOnHGlobal", Ptr, hData, "int", 1, A_PtrSize ? "UPtr*" : "UInt*", pStream)
DllCall("gdiplus\GdipCreateBitmapFromStream", Ptr, pStream, A_PtrSize ? "UPtr*" : "UInt*", pBitmap)
If (A_PtrSize)
%FName%(pStream)
Else
DllCall(NumGet(NumGet(1*pStream)+8), "uint", pStream)
return pBitmap
}
Gdip_DrawRectangle(pGraphics, pPen, x, y, w, h)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawRectangle", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h)
}
Gdip_DrawRoundedRectangle(pGraphics, pPen, x, y, w, h, r)
{
Gdip_SetClipRect(pGraphics, x-r, y-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x+w-r, y-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x-r, y+h-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x+w-r, y+h-r, 2*r, 2*r, 4)
E := Gdip_DrawRectangle(pGraphics, pPen, x, y, w, h)
Gdip_ResetClip(pGraphics)
Gdip_SetClipRect(pGraphics, x-(2*r), y+r, w+(4*r), h-(2*r), 4)
Gdip_SetClipRect(pGraphics, x+r, y-(2*r), w-(2*r), h+(4*r), 4)
Gdip_DrawEllipse(pGraphics, pPen, x, y, 2*r, 2*r)
Gdip_DrawEllipse(pGraphics, pPen, x+w-(2*r), y, 2*r, 2*r)
Gdip_DrawEllipse(pGraphics, pPen, x, y+h-(2*r), 2*r, 2*r)
Gdip_DrawEllipse(pGraphics, pPen, x+w-(2*r), y+h-(2*r), 2*r, 2*r)
Gdip_ResetClip(pGraphics)
return E
}
Gdip_DrawEllipse(pGraphics, pPen, x, y, w, h)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawEllipse", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h)
}
Gdip_DrawBezier(pGraphics, pPen, x1, y1, x2, y2, x3, y3, x4, y4)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawBezier"
, Ptr, pgraphics
, Ptr, pPen
, "float", x1
, "float", y1
, "float", x2
, "float", y2
, "float", x3
, "float", y3
, "float", x4
, "float", y4)
}
Gdip_DrawArc(pGraphics, pPen, x, y, w, h, StartAngle, SweepAngle)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawArc"
, Ptr, pGraphics
, Ptr, pPen
, "float", x
, "float", y
, "float", w
, "float", h
, "float", StartAngle
, "float", SweepAngle)
}
Gdip_DrawPie(pGraphics, pPen, x, y, w, h, StartAngle, SweepAngle)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawPie", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h, "float", StartAngle, "float", SweepAngle)
}
Gdip_DrawLine(pGraphics, pPen, x1, y1, x2, y2)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipDrawLine"
, Ptr, pGraphics
, Ptr, pPen
, "float", x1
, "float", y1
, "float", x2
, "float", y2)
}
Gdip_DrawLines(pGraphics, pPen, Points)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
StringSplit, Points, Points, |
VarSetCapacity(PointF, 8*Points0)
Loop, %Points0%
{
StringSplit, Coord, Points%A_Index%, `,
NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
}
return DllCall("gdiplus\GdipDrawLines", Ptr, pGraphics, Ptr, pPen, Ptr, &PointF, "int", Points0)
}
Gdip_FillRectangle(pGraphics, pBrush, x, y, w, h)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipFillRectangle"
, Ptr, pGraphics
, Ptr, pBrush
, "float", x
, "float", y
, "float", w
, "float", h)
}
Gdip_FillRoundedRectangle(pGraphics, pBrush, x, y, w, h, r)
{
Region := Gdip_GetClipRegion(pGraphics)
Gdip_SetClipRect(pGraphics, x-r, y-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x+w-r, y-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x-r, y+h-r, 2*r, 2*r, 4)
Gdip_SetClipRect(pGraphics, x+w-r, y+h-r, 2*r, 2*r, 4)
E := Gdip_FillRectangle(pGraphics, pBrush, x, y, w, h)
Gdip_SetClipRegion(pGraphics, Region, 0)
Gdip_SetClipRect(pGraphics, x-(2*r), y+r, w+(4*r), h-(2*r), 4)
Gdip_SetClipRect(pGraphics, x+r, y-(2*r), w-(2*r), h+(4*r), 4)
Gdip_FillEllipse(pGraphics, pBrush, x, y, 2*r, 2*r)
Gdip_FillEllipse(pGraphics, pBrush, x+w-(2*r), y, 2*r, 2*r)
Gdip_FillEllipse(pGraphics, pBrush, x, y+h-(2*r), 2*r, 2*r)
Gdip_FillEllipse(pGraphics, pBrush, x+w-(2*r), y+h-(2*r), 2*r, 2*r)
Gdip_SetClipRegion(pGraphics, Region, 0)
Gdip_DeleteRegion(Region)
return E
}
Gdip_FillPolygon(pGraphics, pBrush, Points, FillMode=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
StringSplit, Points, Points, |
VarSetCapacity(PointF, 8*Points0)
Loop, %Points0%
{
StringSplit, Coord, Points%A_Index%, `,
NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
}
return DllCall("gdiplus\GdipFillPolygon", Ptr, pGraphics, Ptr, pBrush, Ptr, &PointF, "int", Points0, "int", FillMode)
}
Gdip_FillPie(pGraphics, pBrush, x, y, w, h, StartAngle, SweepAngle)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipFillPie"
, Ptr, pGraphics
, Ptr, pBrush
, "float", x
, "float", y
, "float", w
, "float", h
, "float", StartAngle
, "float", SweepAngle)
}
Gdip_FillEllipse(pGraphics, pBrush, x, y, w, h)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipFillEllipse", Ptr, pGraphics, Ptr, pBrush, "float", x, "float", y, "float", w, "float", h)
}
Gdip_FillRegion(pGraphics, pBrush, Region)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipFillRegion", Ptr, pGraphics, Ptr, pBrush, Ptr, Region)
}
Gdip_FillPath(pGraphics, pBrush, Path)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipFillPath", Ptr, pGraphics, Ptr, pBrush, Ptr, Path)
}
Gdip_DrawImagePointsRect(pGraphics, pBitmap, Points, sx="", sy="", sw="", sh="", Matrix=1)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
StringSplit, Points, Points, |
VarSetCapacity(PointF, 8*Points0)
Loop, %Points0%
{
StringSplit, Coord, Points%A_Index%, `,
NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
}
if (Matrix&1 = "")
ImageAttr := Gdip_SetImageAttributesColorMatrix(Matrix)
else if (Matrix != 1)
ImageAttr := Gdip_SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")
if (sx = "" && sy = "" && sw = "" && sh = "")
{
sx := 0, sy := 0
sw := Gdip_GetImageWidth(pBitmap)
sh := Gdip_GetImageHeight(pBitmap)
}
E := DllCall("gdiplus\GdipDrawImagePointsRect"
, Ptr, pGraphics
, Ptr, pBitmap
, Ptr, &PointF
, "int", Points0
, "float", sx
, "float", sy
, "float", sw
, "float", sh
, "int", 2
, Ptr, ImageAttr
, Ptr, 0
, Ptr, 0)
if ImageAttr
Gdip_DisposeImageAttributes(ImageAttr)
return E
}
Gdip_DrawImage(pGraphics, pBitmap, dx="", dy="", dw="", dh="", sx="", sy="", sw="", sh="", Matrix=1)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if (Matrix&1 = "")
ImageAttr := Gdip_SetImageAttributesColorMatrix(Matrix)
else if (Matrix != 1)
ImageAttr := Gdip_SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")
if (sx = "" && sy = "" && sw = "" && sh = "")
{
if (dx = "" && dy = "" && dw = "" && dh = "")
{
sx := dx := 0, sy := dy := 0
sw := dw := Gdip_GetImageWidth(pBitmap)
sh := dh := Gdip_GetImageHeight(pBitmap)
}
else
{
sx := sy := 0
sw := Gdip_GetImageWidth(pBitmap)
sh := Gdip_GetImageHeight(pBitmap)
}
}
E := DllCall("gdiplus\GdipDrawImageRectRect"
, Ptr, pGraphics
, Ptr, pBitmap
, "float", dx
, "float", dy
, "float", dw
, "float", dh
, "float", sx
, "float", sy
, "float", sw
, "float", sh
, "int", 2
, Ptr, ImageAttr
, Ptr, 0
, Ptr, 0)
if ImageAttr
Gdip_DisposeImageAttributes(ImageAttr)
return E
}
Gdip_SetImageAttributesColorMatrix(Matrix)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
VarSetCapacity(ColourMatrix, 100, 0)
Matrix := RegExReplace(RegExReplace(Matrix, "^[^\d-\.]+([\d\.])", "$1", "", 1), "[^\d-\.]+", "|")
StringSplit, Matrix, Matrix, |
Loop, 25
{
Matrix := (Matrix%A_Index% != "") ? Matrix%A_Index% : Mod(A_Index-1, 6) ? 0 : 1
NumPut(Matrix, ColourMatrix, (A_Index-1)*4, "float")
}
DllCall("gdiplus\GdipCreateImageAttributes", A_PtrSize ? "UPtr*" : "uint*", ImageAttr)
DllCall("gdiplus\GdipSetImageAttributesColorMatrix", Ptr, ImageAttr, "int", 1, "int", 1, Ptr, &ColourMatrix, Ptr, 0, "int", 0)
return ImageAttr
}
Gdip_GraphicsFromImage(pBitmap)
{
DllCall("gdiplus\GdipGetImageGraphicsContext", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "UInt*", pGraphics)
return pGraphics
}
Gdip_GraphicsFromHDC(hdc)
{
DllCall("gdiplus\GdipCreateFromHDC", A_PtrSize ? "UPtr" : "UInt", hdc, A_PtrSize ? "UPtr*" : "UInt*", pGraphics)
return pGraphics
}
Gdip_GetDC(pGraphics)
{
DllCall("gdiplus\GdipGetDC", A_PtrSize ? "UPtr" : "UInt", pGraphics, A_PtrSize ? "UPtr*" : "UInt*", hdc)
return hdc
}
Gdip_ReleaseDC(pGraphics, hdc)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipReleaseDC", Ptr, pGraphics, Ptr, hdc)
}
Gdip_GraphicsClear(pGraphics, ARGB=0x00ffffff)
{
return DllCall("gdiplus\GdipGraphicsClear", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", ARGB)
}
Gdip_BlurBitmap(pBitmap, Blur)
{
if (Blur > 100) || (Blur < 1)
return -1
sWidth := Gdip_GetImageWidth(pBitmap), sHeight := Gdip_GetImageHeight(pBitmap)
dWidth := sWidth//Blur, dHeight := sHeight//Blur
pBitmap1 := Gdip_CreateBitmap(dWidth, dHeight)
G1 := Gdip_GraphicsFromImage(pBitmap1)
Gdip_SetInterpolationMode(G1, 7)
Gdip_DrawImage(G1, pBitmap, 0, 0, dWidth, dHeight, 0, 0, sWidth, sHeight)
Gdip_DeleteGraphics(G1)
pBitmap2 := Gdip_CreateBitmap(sWidth, sHeight)
G2 := Gdip_GraphicsFromImage(pBitmap2)
Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap1, 0, 0, sWidth, sHeight, 0, 0, dWidth, dHeight)
Gdip_DeleteGraphics(G2)
Gdip_DisposeImage(pBitmap1)
return pBitmap2
}
Gdip_SaveBitmapToFile(pBitmap, sOutput, Quality=75)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
SplitPath, sOutput,,, Extension
if Extension not in BMP,DIB,RLE,JPG,JPEG,JPE,JFIF,GIF,TIF,TIFF,PNG
return -1
Extension := "." Extension
DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", nCount, "uint*", nSize)
VarSetCapacity(ci, nSize)
DllCall("gdiplus\GdipGetImageEncoders", "uint", nCount, "uint", nSize, Ptr, &ci)
if !(nCount && nSize)
return -2
If (A_IsUnicode){
StrGet_Name := "StrGet"
Loop, %nCount%
{
sString := %StrGet_Name%(NumGet(ci, (idx := (48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize), "UTF-16")
if !InStr(sString, "*" Extension)
continue
pCodec := &ci+idx
break
}
} else {
Loop, %nCount%
{
Location := NumGet(ci, 76*(A_Index-1)+44)
nSize := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "uint", 0, "int",  0, "uint", 0, "uint", 0)
VarSetCapacity(sString, nSize)
DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "str", sString, "int", nSize, "uint", 0, "uint", 0)
if !InStr(sString, "*" Extension)
continue
pCodec := &ci+76*(A_Index-1)
break
}
}
if !pCodec
return -3
if (Quality != 75)
{
Quality := (Quality < 0) ? 0 : (Quality > 100) ? 100 : Quality
if Extension in .JPG,.JPEG,.JPE,.JFIF
{
DllCall("gdiplus\GdipGetEncoderParameterListSize", Ptr, pBitmap, Ptr, pCodec, "uint*", nSize)
VarSetCapacity(EncoderParameters, nSize, 0)
DllCall("gdiplus\GdipGetEncoderParameterList", Ptr, pBitmap, Ptr, pCodec, "uint", nSize, Ptr, &EncoderParameters)
Loop, % NumGet(EncoderParameters, "UInt")
{
elem := (24+(A_PtrSize ? A_PtrSize : 4))*(A_Index-1) + 4 + (pad := A_PtrSize = 8 ? 4 : 0)
if (NumGet(EncoderParameters, elem+16, "UInt") = 1) && (NumGet(EncoderParameters, elem+20, "UInt") = 6)
{
p := elem+&EncoderParameters-pad-4
NumPut(Quality, NumGet(NumPut(4, NumPut(1, p+0)+20, "UInt")), "UInt")
break
}
}
}
}
if (!A_IsUnicode)
{
nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sOutput, "int", -1, Ptr, 0, "int", 0)
VarSetCapacity(wOutput, nSize*2)
DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sOutput, "int", -1, Ptr, &wOutput, "int", nSize)
VarSetCapacity(wOutput, -1)
if !VarSetCapacity(wOutput)
return -4
E := DllCall("gdiplus\GdipSaveImageToFile", Ptr, pBitmap, Ptr, &wOutput, Ptr, pCodec, "uint", p ? p : 0)
}
else
E := DllCall("gdiplus\GdipSaveImageToFile", Ptr, pBitmap, Ptr, &sOutput, Ptr, pCodec, "uint", p ? p : 0)
return E ? -5 : 0
}
Gdip_GetPixel(pBitmap, x, y)
{
DllCall("gdiplus\GdipBitmapGetPixel", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", x, "int", y, "uint*", ARGB)
return ARGB
}
Gdip_SetPixel(pBitmap, x, y, ARGB)
{
return DllCall("gdiplus\GdipBitmapSetPixel", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", x, "int", y, "int", ARGB)
}
Gdip_GetImageWidth(pBitmap)
{
DllCall("gdiplus\GdipGetImageWidth", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Width)
return Width
}
Gdip_GetImageHeight(pBitmap)
{
DllCall("gdiplus\GdipGetImageHeight", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Height)
return Height
}
Gdip_GetImageDimensions(pBitmap, ByRef Width, ByRef Height)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
DllCall("gdiplus\GdipGetImageWidth", Ptr, pBitmap, "uint*", Width)
DllCall("gdiplus\GdipGetImageHeight", Ptr, pBitmap, "uint*", Height)
}
Gdip_GetDimensions(pBitmap, ByRef Width, ByRef Height)
{
Gdip_GetImageDimensions(pBitmap, Width, Height)
}
Gdip_GetImagePixelFormat(pBitmap)
{
DllCall("gdiplus\GdipGetImagePixelFormat", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "UInt*", Format)
return Format
}
Gdip_GetDpiX(pGraphics)
{
DllCall("gdiplus\GdipGetDpiX", A_PtrSize ? "UPtr" : "uint", pGraphics, "float*", dpix)
return Round(dpix)
}
Gdip_GetDpiY(pGraphics)
{
DllCall("gdiplus\GdipGetDpiY", A_PtrSize ? "UPtr" : "uint", pGraphics, "float*", dpiy)
return Round(dpiy)
}
Gdip_GetImageHorizontalResolution(pBitmap)
{
DllCall("gdiplus\GdipGetImageHorizontalResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float*", dpix)
return Round(dpix)
}
Gdip_GetImageVerticalResolution(pBitmap)
{
DllCall("gdiplus\GdipGetImageVerticalResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float*", dpiy)
return Round(dpiy)
}
Gdip_BitmapSetResolution(pBitmap, dpix, dpiy)
{
return DllCall("gdiplus\GdipBitmapSetResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float", dpix, "float", dpiy)
}
Gdip_CreateBitmapFromFile(sFile, IconNumber=1, IconSize="")
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
, PtrA := A_PtrSize ? "UPtr*" : "UInt*"
SplitPath, sFile,,, ext
if ext in exe,dll
{
Sizes := IconSize ? IconSize : 256 "|" 128 "|" 64 "|" 48 "|" 32 "|" 16
BufSize := 16 + (2*(A_PtrSize ? A_PtrSize : 4))
VarSetCapacity(buf, BufSize, 0)
Loop, Parse, Sizes, |
{
DllCall("PrivateExtractIcons", "str", sFile, "int", IconNumber-1, "int", A_LoopField, "int", A_LoopField, PtrA, hIcon, PtrA, 0, "uint", 1, "uint", 0)
if !hIcon
continue
if !DllCall("GetIconInfo", Ptr, hIcon, Ptr, &buf)
{
DestroyIcon(hIcon)
continue
}
hbmMask  := NumGet(buf, 12 + ((A_PtrSize ? A_PtrSize : 4) - 4))
hbmColor := NumGet(buf, 12 + ((A_PtrSize ? A_PtrSize : 4) - 4) + (A_PtrSize ? A_PtrSize : 4))
if !(hbmColor && DllCall("GetObject", Ptr, hbmColor, "int", BufSize, Ptr, &buf))
{
DestroyIcon(hIcon)
continue
}
break
}
if !hIcon
return -1
Width := NumGet(buf, 4, "int"), Height := NumGet(buf, 8, "int")
hbm := CreateDIBSection(Width, -Height), hdc := CreateCompatibleDC(), obm := SelectObject(hdc, hbm)
if !DllCall("DrawIconEx", Ptr, hdc, "int", 0, "int", 0, Ptr, hIcon, "uint", Width, "uint", Height, "uint", 0, Ptr, 0, "uint", 3)
{
DestroyIcon(hIcon)
return -2
}
VarSetCapacity(dib, 104)
DllCall("GetObject", Ptr, hbm, "int", A_PtrSize = 8 ? 104 : 84, Ptr, &dib)
Stride := NumGet(dib, 12, "Int"), Bits := NumGet(dib, 20 + (A_PtrSize = 8 ? 4 : 0))
DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", Width, "int", Height, "int", Stride, "int", 0x26200A, Ptr, Bits, PtrA, pBitmapOld)
pBitmap := Gdip_CreateBitmap(Width, Height)
G := Gdip_GraphicsFromImage(pBitmap)
, Gdip_DrawImage(G, pBitmapOld, 0, 0, Width, Height, 0, 0, Width, Height)
SelectObject(hdc, obm), DeleteObject(hbm), DeleteDC(hdc)
Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmapOld)
DestroyIcon(hIcon)
}
else
{
if (!A_IsUnicode)
{
VarSetCapacity(wFile, 1024)
DllCall("kernel32\MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sFile, "int", -1, Ptr, &wFile, "int", 512)
DllCall("gdiplus\GdipCreateBitmapFromFile", Ptr, &wFile, PtrA, pBitmap)
}
else
DllCall("gdiplus\GdipCreateBitmapFromFile", Ptr, &sFile, PtrA, pBitmap)
}
return pBitmap
}
Gdip_CreateBitmapFromHBITMAP(hBitmap, Palette=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", Ptr, hBitmap, Ptr, Palette, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
return pBitmap
}
Gdip_CreateHBITMAPFromBitmap(pBitmap, Background=0xffffffff)
{
DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hbm, "int", Background)
return hbm
}
Gdip_CreateBitmapFromHICON(hIcon)
{
DllCall("gdiplus\GdipCreateBitmapFromHICON", A_PtrSize ? "UPtr" : "UInt", hIcon, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
return pBitmap
}
Gdip_CreateHICONFromBitmap(pBitmap)
{
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
return hIcon
}
Gdip_CreateBitmap(Width, Height, Format=0x26200A)
{
DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", Width, "int", Height, "int", 0, "int", Format, A_PtrSize ? "UPtr" : "UInt", 0, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
Return pBitmap
}
Gdip_CreateBitmapFromClipboard()
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if !DllCall("OpenClipboard", Ptr, 0)
return -1
if !DllCall("IsClipboardFormatAvailable", "uint", 8)
return -2
if !hBitmap := DllCall("GetClipboardData", "uint", 2, Ptr)
return -3
if !pBitmap := Gdip_CreateBitmapFromHBITMAP(hBitmap)
return -4
if !DllCall("CloseClipboard")
return -5
DeleteObject(hBitmap)
return pBitmap
}
Gdip_SetBitmapToClipboard(pBitmap)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
off1 := A_PtrSize = 8 ? 52 : 44, off2 := A_PtrSize = 8 ? 32 : 24
hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
DllCall("GetObject", Ptr, hBitmap, "int", VarSetCapacity(oi, A_PtrSize = 8 ? 104 : 84, 0), Ptr, &oi)
hdib := DllCall("GlobalAlloc", "uint", 2, Ptr, 40+NumGet(oi, off1, "UInt"), Ptr)
pdib := DllCall("GlobalLock", Ptr, hdib, Ptr)
DllCall("RtlMoveMemory", Ptr, pdib, Ptr, &oi+off2, Ptr, 40)
DllCall("RtlMoveMemory", Ptr, pdib+40, Ptr, NumGet(oi, off2 - (A_PtrSize ? A_PtrSize : 4), Ptr), Ptr, NumGet(oi, off1, "UInt"))
DllCall("GlobalUnlock", Ptr, hdib)
DllCall("DeleteObject", Ptr, hBitmap)
DllCall("OpenClipboard", Ptr, 0)
DllCall("EmptyClipboard")
DllCall("SetClipboardData", "uint", 8, Ptr, hdib)
DllCall("CloseClipboard")
}
Gdip_CloneBitmapArea(pBitmap, x, y, w, h, Format=0x26200A)
{
DllCall("gdiplus\GdipCloneBitmapArea"
, "float", x
, "float", y
, "float", w
, "float", h
, "int", Format
, A_PtrSize ? "UPtr" : "UInt", pBitmap
, A_PtrSize ? "UPtr*" : "UInt*", pBitmapDest)
return pBitmapDest
}
Gdip_CreatePen(ARGB, w)
{
DllCall("gdiplus\GdipCreatePen1", "UInt", ARGB, "float", w, "int", 2, A_PtrSize ? "UPtr*" : "UInt*", pPen)
return pPen
}
Gdip_CreatePenFromBrush(pBrush, w)
{
DllCall("gdiplus\GdipCreatePen2", A_PtrSize ? "UPtr" : "UInt", pBrush, "float", w, "int", 2, A_PtrSize ? "UPtr*" : "UInt*", pPen)
return pPen
}
Gdip_BrushCreateSolid(ARGB=0xff000000)
{
DllCall("gdiplus\GdipCreateSolidFill", "UInt", ARGB, A_PtrSize ? "UPtr*" : "UInt*", pBrush)
return pBrush
}
Gdip_BrushCreateHatch(ARGBfront, ARGBback, HatchStyle=0)
{
DllCall("gdiplus\GdipCreateHatchBrush", "int", HatchStyle, "UInt", ARGBfront, "UInt", ARGBback, A_PtrSize ? "UPtr*" : "UInt*", pBrush)
return pBrush
}
Gdip_CreateTextureBrush(pBitmap, WrapMode=1, x=0, y=0, w="", h="")
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
, PtrA := A_PtrSize ? "UPtr*" : "UInt*"
if !(w && h)
DllCall("gdiplus\GdipCreateTexture", Ptr, pBitmap, "int", WrapMode, PtrA, pBrush)
else
DllCall("gdiplus\GdipCreateTexture2", Ptr, pBitmap, "int", WrapMode, "float", x, "float", y, "float", w, "float", h, PtrA, pBrush)
return pBrush
}
Gdip_CreateLineBrush(x1, y1, x2, y2, ARGB1, ARGB2, WrapMode=1)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
CreatePointF(PointF1, x1, y1), CreatePointF(PointF2, x2, y2)
DllCall("gdiplus\GdipCreateLineBrush", Ptr, &PointF1, Ptr, &PointF2, "Uint", ARGB1, "Uint", ARGB2, "int", WrapMode, A_PtrSize ? "UPtr*" : "UInt*", LGpBrush)
return LGpBrush
}
Gdip_CreateLineBrushFromRect(x, y, w, h, ARGB1, ARGB2, LinearGradientMode=1, WrapMode=1)
{
CreateRectF(RectF, x, y, w, h)
DllCall("gdiplus\GdipCreateLineBrushFromRect", A_PtrSize ? "UPtr" : "UInt", &RectF, "int", ARGB1, "int", ARGB2, "int", LinearGradientMode, "int", WrapMode, A_PtrSize ? "UPtr*" : "UInt*", LGpBrush)
return LGpBrush
}
Gdip_CloneBrush(pBrush)
{
DllCall("gdiplus\GdipCloneBrush", A_PtrSize ? "UPtr" : "UInt", pBrush, A_PtrSize ? "UPtr*" : "UInt*", pBrushClone)
return pBrushClone
}
Gdip_DeletePen(pPen)
{
return DllCall("gdiplus\GdipDeletePen", A_PtrSize ? "UPtr" : "UInt", pPen)
}
Gdip_DeleteBrush(pBrush)
{
return DllCall("gdiplus\GdipDeleteBrush", A_PtrSize ? "UPtr" : "UInt", pBrush)
}
Gdip_DisposeImage(pBitmap)
{
return DllCall("gdiplus\GdipDisposeImage", A_PtrSize ? "UPtr" : "UInt", pBitmap)
}
Gdip_DeleteGraphics(pGraphics)
{
return DllCall("gdiplus\GdipDeleteGraphics", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}
Gdip_DisposeImageAttributes(ImageAttr)
{
return DllCall("gdiplus\GdipDisposeImageAttributes", A_PtrSize ? "UPtr" : "UInt", ImageAttr)
}
Gdip_DeleteFont(hFont)
{
return DllCall("gdiplus\GdipDeleteFont", A_PtrSize ? "UPtr" : "UInt", hFont)
}
Gdip_DeleteStringFormat(hFormat)
{
return DllCall("gdiplus\GdipDeleteStringFormat", A_PtrSize ? "UPtr" : "UInt", hFormat)
}
Gdip_DeleteFontFamily(hFamily)
{
return DllCall("gdiplus\GdipDeleteFontFamily", A_PtrSize ? "UPtr" : "UInt", hFamily)
}
Gdip_DeleteMatrix(Matrix)
{
return DllCall("gdiplus\GdipDeleteMatrix", A_PtrSize ? "UPtr" : "UInt", Matrix)
}
Gdip_TextToGraphics(pGraphics, Text, Options, Font="Arial", Width="", Height="", Measure=0)
{
IWidth := Width, IHeight:= Height
RegExMatch(Options, "i)X([\-\d\.]+)(p*)", xpos)
RegExMatch(Options, "i)Y([\-\d\.]+)(p*)", ypos)
RegExMatch(Options, "i)W([\-\d\.]+)(p*)", Width)
RegExMatch(Options, "i)H([\-\d\.]+)(p*)", Height)
RegExMatch(Options, "i)C(?!(entre|enter))([a-f\d]+)", Colour)
RegExMatch(Options, "i)Top|Up|Bottom|Down|vCentre|vCenter", vPos)
RegExMatch(Options, "i)NoWrap", NoWrap)
RegExMatch(Options, "i)R(\d)", Rendering)
RegExMatch(Options, "i)S(\d+)(p*)", Size)
if !Gdip_DeleteBrush(Gdip_CloneBrush(Colour2))
PassBrush := 1, pBrush := Colour2
if !(IWidth && IHeight) && (xpos2 || ypos2 || Width2 || Height2 || Size2)
return -1
Style := 0, Styles := "Regular|Bold|Italic|BoldItalic|Underline|Strikeout"
Loop, Parse, Styles, |
{
if RegExMatch(Options, "\b" A_loopField)
Style |= (A_LoopField != "StrikeOut") ? (A_Index-1) : 8
}
Align := 0, Alignments := "Near|Left|Centre|Center|Far|Right"
Loop, Parse, Alignments, |
{
if RegExMatch(Options, "\b" A_loopField)
Align |= A_Index//2.1
}
xpos := (xpos1 != "") ? xpos2 ? IWidth*(xpos1/100) : xpos1 : 0
ypos := (ypos1 != "") ? ypos2 ? IHeight*(ypos1/100) : ypos1 : 0
Width := Width1 ? Width2 ? IWidth*(Width1/100) : Width1 : IWidth
Height := Height1 ? Height2 ? IHeight*(Height1/100) : Height1 : IHeight
if !PassBrush
Colour := "0x" (Colour2 ? Colour2 : "ff000000")
Rendering := ((Rendering1 >= 0) && (Rendering1 <= 5)) ? Rendering1 : 4
Size := (Size1 > 0) ? Size2 ? IHeight*(Size1/100) : Size1 : 12
hFamily := Gdip_FontFamilyCreate(Font)
hFont := Gdip_FontCreate(hFamily, Size, Style)
FormatStyle := NoWrap ? 0x4000 | 0x1000 : 0x4000
hFormat := Gdip_StringFormatCreate(FormatStyle)
pBrush := PassBrush ? pBrush : Gdip_BrushCreateSolid(Colour)
if !(hFamily && hFont && hFormat && pBrush && pGraphics)
return !pGraphics ? -2 : !hFamily ? -3 : !hFont ? -4 : !hFormat ? -5 : !pBrush ? -6 : 0
CreateRectF(RC, xpos, ypos, Width, Height)
Gdip_SetStringFormatAlign(hFormat, Align)
Gdip_SetTextRenderingHint(pGraphics, Rendering)
ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)
if vPos
{
StringSplit, ReturnRC, ReturnRC, |
if (vPos = "vCentre") || (vPos = "vCenter")
ypos += (Height-ReturnRC4)//2
else if (vPos = "Top") || (vPos = "Up")
ypos := 0
else if (vPos = "Bottom") || (vPos = "Down")
ypos := Height-ReturnRC4
CreateRectF(RC, xpos, ypos, Width, ReturnRC4)
ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)
}
if !Measure
E := Gdip_DrawString(pGraphics, Text, hFont, hFormat, pBrush, RC)
if !PassBrush
Gdip_DeleteBrush(pBrush)
Gdip_DeleteStringFormat(hFormat)
Gdip_DeleteFont(hFont)
Gdip_DeleteFontFamily(hFamily)
return E ? E : ReturnRC
}
Gdip_DrawString(pGraphics, sString, hFont, hFormat, pBrush, ByRef RectF)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if (!A_IsUnicode)
{
nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, 0, "int", 0)
VarSetCapacity(wString, nSize*2)
DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, &wString, "int", nSize)
}
return DllCall("gdiplus\GdipDrawString"
, Ptr, pGraphics
, Ptr, A_IsUnicode ? &sString : &wString
, "int", -1
, Ptr, hFont
, Ptr, &RectF
, Ptr, hFormat
, Ptr, pBrush)
}
Gdip_MeasureString(pGraphics, sString, hFont, hFormat, ByRef RectF)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
VarSetCapacity(RC, 16)
if !A_IsUnicode
{
nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, "uint", 0, "int", 0)
VarSetCapacity(wString, nSize*2)
DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, &wString, "int", nSize)
}
DllCall("gdiplus\GdipMeasureString"
, Ptr, pGraphics
, Ptr, A_IsUnicode ? &sString : &wString
, "int", -1
, Ptr, hFont
, Ptr, &RectF
, Ptr, hFormat
, Ptr, &RC
, "uint*", Chars
, "uint*", Lines)
return &RC ? NumGet(RC, 0, "float") "|" NumGet(RC, 4, "float") "|" NumGet(RC, 8, "float") "|" NumGet(RC, 12, "float") "|" Chars "|" Lines : 0
}
Gdip_SetStringFormatAlign(hFormat, Align)
{
return DllCall("gdiplus\GdipSetStringFormatAlign", A_PtrSize ? "UPtr" : "UInt", hFormat, "int", Align)
}
Gdip_StringFormatCreate(Format=0, Lang=0)
{
DllCall("gdiplus\GdipCreateStringFormat", "int", Format, "int", Lang, A_PtrSize ? "UPtr*" : "UInt*", hFormat)
return hFormat
}
Gdip_FontCreate(hFamily, Size, Style=0)
{
DllCall("gdiplus\GdipCreateFont", A_PtrSize ? "UPtr" : "UInt", hFamily, "float", Size, "int", Style, "int", 0, A_PtrSize ? "UPtr*" : "UInt*", hFont)
return hFont
}
Gdip_FontFamilyCreate(Font)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if (!A_IsUnicode)
{
nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &Font, "int", -1, "uint", 0, "int", 0)
VarSetCapacity(wFont, nSize*2)
DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &Font, "int", -1, Ptr, &wFont, "int", nSize)
}
DllCall("gdiplus\GdipCreateFontFamilyFromName"
, Ptr, A_IsUnicode ? &Font : &wFont
, "uint", 0
, A_PtrSize ? "UPtr*" : "UInt*", hFamily)
return hFamily
}
Gdip_CreateAffineMatrix(m11, m12, m21, m22, x, y)
{
DllCall("gdiplus\GdipCreateMatrix2", "float", m11, "float", m12, "float", m21, "float", m22, "float", x, "float", y, A_PtrSize ? "UPtr*" : "UInt*", Matrix)
return Matrix
}
Gdip_CreateMatrix()
{
DllCall("gdiplus\GdipCreateMatrix", A_PtrSize ? "UPtr*" : "UInt*", Matrix)
return Matrix
}
Gdip_CreatePath(BrushMode=0)
{
DllCall("gdiplus\GdipCreatePath", "int", BrushMode, A_PtrSize ? "UPtr*" : "UInt*", Path)
return Path
}
Gdip_AddPathEllipse(Path, x, y, w, h)
{
return DllCall("gdiplus\GdipAddPathEllipse", A_PtrSize ? "UPtr" : "UInt", Path, "float", x, "float", y, "float", w, "float", h)
}
Gdip_AddPathPolygon(Path, Points)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
StringSplit, Points, Points, |
VarSetCapacity(PointF, 8*Points0)
Loop, %Points0%
{
StringSplit, Coord, Points%A_Index%, `,
NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
}
return DllCall("gdiplus\GdipAddPathPolygon", Ptr, Path, Ptr, &PointF, "int", Points0)
}
Gdip_DeletePath(Path)
{
return DllCall("gdiplus\GdipDeletePath", A_PtrSize ? "UPtr" : "UInt", Path)
}
Gdip_SetTextRenderingHint(pGraphics, RenderingHint)
{
return DllCall("gdiplus\GdipSetTextRenderingHint", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", RenderingHint)
}
Gdip_SetInterpolationMode(pGraphics, InterpolationMode)
{
return DllCall("gdiplus\GdipSetInterpolationMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", InterpolationMode)
}
Gdip_SetSmoothingMode(pGraphics, SmoothingMode)
{
return DllCall("gdiplus\GdipSetSmoothingMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", SmoothingMode)
}
Gdip_SetCompositingMode(pGraphics, CompositingMode=0)
{
return DllCall("gdiplus\GdipSetCompositingMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", CompositingMode)
}
Gdip_Startup()
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
if !DllCall("GetModuleHandle", "str", "gdiplus", Ptr)
DllCall("LoadLibrary", "str", "gdiplus")
VarSetCapacity(si, A_PtrSize = 8 ? 24 : 16, 0), si := Chr(1)
DllCall("gdiplus\GdiplusStartup", A_PtrSize ? "UPtr*" : "uint*", pToken, Ptr, &si, Ptr, 0)
return pToken
}
Gdip_Shutdown(pToken)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
DllCall("gdiplus\GdiplusShutdown", Ptr, pToken)
if hModule := DllCall("GetModuleHandle", "str", "gdiplus", Ptr)
DllCall("FreeLibrary", Ptr, hModule)
return 0
}
Gdip_RotateWorldTransform(pGraphics, Angle, MatrixOrder=0)
{
return DllCall("gdiplus\GdipRotateWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", Angle, "int", MatrixOrder)
}
Gdip_ScaleWorldTransform(pGraphics, x, y, MatrixOrder=0)
{
return DllCall("gdiplus\GdipScaleWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "int", MatrixOrder)
}
Gdip_TranslateWorldTransform(pGraphics, x, y, MatrixOrder=0)
{
return DllCall("gdiplus\GdipTranslateWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "int", MatrixOrder)
}
Gdip_ResetWorldTransform(pGraphics)
{
return DllCall("gdiplus\GdipResetWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}
Gdip_GetRotatedTranslation(Width, Height, Angle, ByRef xTranslation, ByRef yTranslation)
{
pi := 3.14159, TAngle := Angle*(pi/180)
Bound := (Angle >= 0) ? Mod(Angle, 360) : 360-Mod(-Angle, -360)
if ((Bound >= 0) && (Bound <= 90))
xTranslation := Height*Sin(TAngle), yTranslation := 0
else if ((Bound > 90) && (Bound <= 180))
xTranslation := (Height*Sin(TAngle))-(Width*Cos(TAngle)), yTranslation := -Height*Cos(TAngle)
else if ((Bound > 180) && (Bound <= 270))
xTranslation := -(Width*Cos(TAngle)), yTranslation := -(Height*Cos(TAngle))-(Width*Sin(TAngle))
else if ((Bound > 270) && (Bound <= 360))
xTranslation := 0, yTranslation := -Width*Sin(TAngle)
}
Gdip_GetRotatedDimensions(Width, Height, Angle, ByRef RWidth, ByRef RHeight)
{
pi := 3.14159, TAngle := Angle*(pi/180)
if !(Width && Height)
return -1
RWidth := Ceil(Abs(Width*Cos(TAngle))+Abs(Height*Sin(TAngle)))
RHeight := Ceil(Abs(Width*Sin(TAngle))+Abs(Height*Cos(Tangle)))
}
Gdip_ImageRotateFlip(pBitmap, RotateFlipType=1)
{
return DllCall("gdiplus\GdipImageRotateFlip", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", RotateFlipType)
}
Gdip_SetClipRect(pGraphics, x, y, w, h, CombineMode=0)
{
return DllCall("gdiplus\GdipSetClipRect",  A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "float", w, "float", h, "int", CombineMode)
}
Gdip_SetClipPath(pGraphics, Path, CombineMode=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipSetClipPath", Ptr, pGraphics, Ptr, Path, "int", CombineMode)
}
Gdip_ResetClip(pGraphics)
{
return DllCall("gdiplus\GdipResetClip", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}
Gdip_GetClipRegion(pGraphics)
{
Region := Gdip_CreateRegion()
DllCall("gdiplus\GdipGetClip", A_PtrSize ? "UPtr" : "UInt", pGraphics, "UInt*", Region)
return Region
}
Gdip_SetClipRegion(pGraphics, Region, CombineMode=0)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("gdiplus\GdipSetClipRegion", Ptr, pGraphics, Ptr, Region, "int", CombineMode)
}
Gdip_CreateRegion()
{
DllCall("gdiplus\GdipCreateRegion", "UInt*", Region)
return Region
}
Gdip_DeleteRegion(Region)
{
return DllCall("gdiplus\GdipDeleteRegion", A_PtrSize ? "UPtr" : "UInt", Region)
}
Gdip_LockBits(pBitmap, x, y, w, h, ByRef Stride, ByRef Scan0, ByRef BitmapData, LockMode = 3, PixelFormat = 0x26200a)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
CreateRect(Rect, x, y, w, h)
VarSetCapacity(BitmapData, 16+2*(A_PtrSize ? A_PtrSize : 4), 0)
E := DllCall("Gdiplus\GdipBitmapLockBits", Ptr, pBitmap, Ptr, &Rect, "uint", LockMode, "int", PixelFormat, Ptr, &BitmapData)
Stride := NumGet(BitmapData, 8, "Int")
Scan0 := NumGet(BitmapData, 16, Ptr)
return E
}
Gdip_UnlockBits(pBitmap, ByRef BitmapData)
{
Ptr := A_PtrSize ? "UPtr" : "UInt"
return DllCall("Gdiplus\GdipBitmapUnlockBits", Ptr, pBitmap, Ptr, &BitmapData)
}
Gdip_SetLockBitPixel(ARGB, Scan0, x, y, Stride)
{
Numput(ARGB, Scan0+0, (x*4)+(y*Stride), "UInt")
}
Gdip_GetLockBitPixel(Scan0, x, y, Stride)
{
return NumGet(Scan0+0, (x*4)+(y*Stride), "UInt")
}
Gdip_PixelateBitmap(pBitmap, ByRef pBitmapOut, BlockSize)
{
static PixelateBitmap
Ptr := A_PtrSize ? "UPtr" : "UInt"
if (!PixelateBitmap)
{
if A_PtrSize != 8
MCode_PixelateBitmap =
		(LTrim Join
		558BEC83EC3C8B4514538B5D1C99F7FB56578BC88955EC894DD885C90F8E830200008B451099F7FB8365DC008365E000894DC88955F08945E833FF897DD4
		397DE80F8E160100008BCB0FAFCB894DCC33C08945F88945FC89451C8945143BD87E608B45088D50028BC82BCA8BF02BF2418945F48B45E02955F4894DC4
		8D0CB80FAFCB03CA895DD08BD1895DE40FB64416030145140FB60201451C8B45C40FB604100145FC8B45F40FB604020145F883C204FF4DE475D6034D18FF
		4DD075C98B4DCC8B451499F7F98945148B451C99F7F989451C8B45FC99F7F98945FC8B45F899F7F98945F885DB7E648B450C8D50028BC82BCA83C103894D
		C48BC82BCA41894DF48B4DD48945E48B45E02955E48D0C880FAFCB03CA895DD08BD18BF38A45148B7DC48804178A451C8B7DF488028A45FC8804178A45F8
		8B7DE488043A83C2044E75DA034D18FF4DD075CE8B4DCC8B7DD447897DD43B7DE80F8CF2FEFFFF837DF0000F842C01000033C08945F88945FC89451C8945
		148945E43BD87E65837DF0007E578B4DDC034DE48B75E80FAF4D180FAFF38B45088D500203CA8D0CB18BF08BF88945F48B45F02BF22BFA2955F48945CC0F
		B6440E030145140FB60101451C0FB6440F010145FC8B45F40FB604010145F883C104FF4DCC75D8FF45E4395DE47C9B8B4DF00FAFCB85C9740B8B451499F7
		F9894514EB048365140033F63BCE740B8B451C99F7F989451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB
		038975F88975E43BDE7E5A837DF0007E4C8B4DDC034DE48B75E80FAF4D180FAFF38B450C8D500203CA8D0CB18BF08BF82BF22BFA2BC28B55F08955CC8A55
		1488540E038A551C88118A55FC88540F018A55F888140183C104FF4DCC75DFFF45E4395DE47CA68B45180145E0015DDCFF4DC80F8594FDFFFF8B451099F7
		FB8955F08945E885C00F8E450100008B45EC0FAFC38365DC008945D48B45E88945CC33C08945F88945FC89451C8945148945103945EC7E6085DB7E518B4D
		D88B45080FAFCB034D108D50020FAF4D18034DDC8BF08BF88945F403CA2BF22BFA2955F4895DC80FB6440E030145140FB60101451C0FB6440F010145FC8B
		45F40FB604080145F883C104FF4DC875D8FF45108B45103B45EC7CA08B4DD485C9740B8B451499F7F9894514EB048365140033F63BCE740B8B451C99F7F9
		89451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB038975F88975103975EC7E5585DB7E468B4DD88B450C
		0FAFCB034D108D50020FAF4D18034DDC8BF08BF803CA2BF22BFA2BC2895DC88A551488540E038A551C88118A55FC88540F018A55F888140183C104FF4DC8
		75DFFF45108B45103B45EC7CAB8BC3C1E0020145DCFF4DCC0F85CEFEFFFF8B4DEC33C08945F88945FC89451C8945148945103BC87E6C3945F07E5C8B4DD8
		8B75E80FAFCB034D100FAFF30FAF4D188B45088D500203CA8D0CB18BF08BF88945F48B45F02BF22BFA2955F48945C80FB6440E030145140FB60101451C0F
		B6440F010145FC8B45F40FB604010145F883C104FF4DC875D833C0FF45108B4DEC394D107C940FAF4DF03BC874068B451499F7F933F68945143BCE740B8B
		451C99F7F989451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB038975F88975083975EC7E63EB0233F639
		75F07E4F8B4DD88B75E80FAFCB034D080FAFF30FAF4D188B450C8D500203CA8D0CB18BF08BF82BF22BFA2BC28B55F08955108A551488540E038A551C8811
		8A55FC88540F018A55F888140883C104FF4D1075DFFF45088B45083B45EC7C9F5F5E33C05BC9C21800
)
else
MCode_PixelateBitmap =
		(LTrim Join
		4489442418488954241048894C24085355565741544155415641574883EC28418BC1448B8C24980000004C8BDA99488BD941F7F9448BD0448BFA8954240C
		448994248800000085C00F8E9D020000418BC04533E4458BF299448924244C8954241041F7F933C9898C24980000008BEA89542404448BE889442408EB05
		4C8B5C24784585ED0F8E1A010000458BF1418BFD48897C2418450FAFF14533D233F633ED4533E44533ED4585C97E5B4C63BC2490000000418D040A410FAF
		C148984C8D441802498BD9498BD04D8BD90FB642010FB64AFF4403E80FB60203E90FB64AFE4883C2044403E003F149FFCB75DE4D03C748FFCB75D0488B7C
		24188B8C24980000004C8B5C2478418BC59941F7FE448BE8418BC49941F7FE448BE08BC59941F7FE8BE88BC69941F7FE8BF04585C97E4048639C24900000
		004103CA4D8BC1410FAFC94863C94A8D541902488BCA498BC144886901448821408869FF408871FE4883C10448FFC875E84803D349FFC875DA8B8C249800
		0000488B5C24704C8B5C24784183C20448FFCF48897C24180F850AFFFFFF8B6C2404448B2424448B6C24084C8B74241085ED0F840A01000033FF33DB4533
		DB4533D24533C04585C97E53488B74247085ED7E42438D0C04418BC50FAF8C2490000000410FAFC18D04814863C8488D5431028BCD0FB642014403D00FB6
		024883C2044403D80FB642FB03D80FB642FA03F848FFC975DE41FFC0453BC17CB28BCD410FAFC985C9740A418BC299F7F98BF0EB0233F685C9740B418BC3
		99F7F9448BD8EB034533DB85C9740A8BC399F7F9448BD0EB034533D285C9740A8BC799F7F9448BC0EB034533C033D24585C97E4D4C8B74247885ED7E3841
		8D0C14418BC50FAF8C2490000000410FAFC18D04814863C84A8D4431028BCD40887001448818448850FF448840FE4883C00448FFC975E8FFC2413BD17CBD
		4C8B7424108B8C2498000000038C2490000000488B5C24704503E149FFCE44892424898C24980000004C897424100F859EFDFFFF448B7C240C448B842480
		000000418BC09941F7F98BE8448BEA89942498000000896C240C85C00F8E3B010000448BAC2488000000418BCF448BF5410FAFC9898C248000000033FF33
		ED33F64533DB4533D24533C04585FF7E524585C97E40418BC5410FAFC14103C00FAF84249000000003C74898488D541802498BD90FB642014403D00FB602
		4883C2044403D80FB642FB03F00FB642FA03E848FFCB75DE488B5C247041FFC0453BC77CAE85C9740B418BC299F7F9448BE0EB034533E485C9740A418BC3
		99F7F98BD8EB0233DB85C9740A8BC699F7F9448BD8EB034533DB85C9740A8BC599F7F9448BD0EB034533D24533C04585FF7E4E488B4C24784585C97E3541
		8BC5410FAFC14103C00FAF84249000000003C74898488D540802498BC144886201881A44885AFF448852FE4883C20448FFC875E941FFC0453BC77CBE8B8C
		2480000000488B5C2470418BC1C1E00203F849FFCE0F85ECFEFFFF448BAC24980000008B6C240C448BA4248800000033FF33DB4533DB4533D24533C04585
		FF7E5A488B7424704585ED7E48418BCC8BC5410FAFC94103C80FAF8C2490000000410FAFC18D04814863C8488D543102418BCD0FB642014403D00FB60248
		83C2044403D80FB642FB03D80FB642FA03F848FFC975DE41FFC0453BC77CAB418BCF410FAFCD85C9740A418BC299F7F98BF0EB0233F685C9740B418BC399
		F7F9448BD8EB034533DB85C9740A8BC399F7F9448BD0EB034533D285C9740A8BC799F7F9448BC0EB034533C033D24585FF7E4E4585ED7E42418BCC8BC541
		0FAFC903CA0FAF8C2490000000410FAFC18D04814863C8488B442478488D440102418BCD40887001448818448850FF448840FE4883C00448FFC975E8FFC2
		413BD77CB233C04883C428415F415E415D415C5F5E5D5BC3
)
VarSetCapacity(PixelateBitmap, StrLen(MCode_PixelateBitmap)//2)
Loop % StrLen(MCode_PixelateBitmap)//2
NumPut("0x" SubStr(MCode_PixelateBitmap, (2*A_Index)-1, 2), PixelateBitmap, A_Index-1, "UChar")
DllCall("VirtualProtect", Ptr, &PixelateBitmap, Ptr, VarSetCapacity(PixelateBitmap), "uint", 0x40, A_PtrSize ? "UPtr*" : "UInt*", 0)
}
Gdip_GetImageDimensions(pBitmap, Width, Height)
if (Width != Gdip_GetImageWidth(pBitmapOut) || Height != Gdip_GetImageHeight(pBitmapOut))
return -1
if (BlockSize > Width || BlockSize > Height)
return -2
E1 := Gdip_LockBits(pBitmap, 0, 0, Width, Height, Stride1, Scan01, BitmapData1)
E2 := Gdip_LockBits(pBitmapOut, 0, 0, Width, Height, Stride2, Scan02, BitmapData2)
if (E1 || E2)
return -3
E := DllCall(&PixelateBitmap, Ptr, Scan01, Ptr, Scan02, "int", Width, "int", Height, "int", Stride1, "int", BlockSize)
Gdip_UnlockBits(pBitmap, BitmapData1), Gdip_UnlockBits(pBitmapOut, BitmapData2)
return 0
}
Gdip_ToARGB(A, R, G, B)
{
return (A << 24) | (R << 16) | (G << 8) | B
}
Gdip_FromARGB(ARGB, ByRef A, ByRef R, ByRef G, ByRef B)
{
A := (0xff000000 & ARGB) >> 24
R := (0x00ff0000 & ARGB) >> 16
G := (0x0000ff00 & ARGB) >> 8
B := 0x000000ff & ARGB
}
Gdip_AFromARGB(ARGB)
{
return (0xff000000 & ARGB) >> 24
}
Gdip_RFromARGB(ARGB)
{
return (0x00ff0000 & ARGB) >> 16
}
Gdip_GFromARGB(ARGB)
{
return (0x0000ff00 & ARGB) >> 8
}
Gdip_BFromARGB(ARGB)
{
return 0x000000ff & ARGB
}
StrGetB(Address, Length=-1, Encoding=0)
{
if Length is not integer
Encoding := Length,  Length := -1
if (Address+0 < 1024)
return
if Encoding = UTF-16
Encoding = 1200
else if Encoding = UTF-8
Encoding = 65001
else if SubStr(Encoding,1,2)="CP"
Encoding := SubStr(Encoding,3)
if !Encoding
{
if (Length == -1)
Length := DllCall("lstrlen", "uint", Address)
VarSetCapacity(String, Length)
DllCall("lstrcpyn", "str", String, "uint", Address, "int", Length + 1)
}
else if Encoding = 1200
{
char_count := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0x400, "uint", Address, "int", Length, "uint", 0, "uint", 0, "uint", 0, "uint", 0)
VarSetCapacity(String, char_count)
DllCall("WideCharToMultiByte", "uint", 0, "uint", 0x400, "uint", Address, "int", Length, "str", String, "int", char_count, "uint", 0, "uint", 0)
}
else if Encoding is integer
{
char_count := DllCall("MultiByteToWideChar", "uint", Encoding, "uint", 0, "uint", Address, "int", Length, "uint", 0, "int", 0)
VarSetCapacity(String, char_count * 2)
char_count := DllCall("MultiByteToWideChar", "uint", Encoding, "uint", 0, "uint", Address, "int", Length, "uint", &String, "int", char_count * 2)
String := StrGetB(&String, char_count, 1200)
}
return String
}
Gdip_ImageSearch(pBitmapHaystack,pBitmapNeedle,ByRef OutputList=""
,OuterX1=0,OuterY1=0,OuterX2=0,OuterY2=0,Variation=0,Trans=""
,SearchDirection=1,Instances=1,LineDelim="`n",CoordDelim=",") {
If !( pBitmapHaystack && pBitmapNeedle )
Return -1001
If Variation not between 0 and 255
return -1002
If ( ( OuterX1 < 0 ) || ( OuterY1 < 0 ) )
return -1003
If SearchDirection not between 1 and 8
SearchDirection := 1
If ( Instances < 0 )
Instances := 0
Gdip_GetImageDimensions(pBitmapHaystack,hWidth,hHeight)
If Gdip_LockBits(pBitmapHaystack,0,0,hWidth,hHeight,hStride,hScan,hBitmapData,1)
OR !(hWidth := NumGet(hBitmapData,0))
OR !(hHeight := NumGet(hBitmapData,4))
Return -1004
Gdip_GetImageDimensions(pBitmapNeedle,nWidth,nHeight)
If Trans between 0 and 0xFFFFFF
{
pOriginalBmpNeedle := pBitmapNeedle
pBitmapNeedle := Gdip_CloneBitmapArea(pOriginalBmpNeedle,0,0,nWidth,nHeight)
Gdip_SetBitmapTransColor(pBitmapNeedle,Trans)
DumpCurrentNeedle := true
}
If Gdip_LockBits(pBitmapNeedle,0,0,nWidth,nHeight,nStride,nScan,nBitmapData)
OR !(nWidth := NumGet(nBitmapData,0))
OR !(nHeight := NumGet(nBitmapData,4))
{
If ( DumpCurrentNeedle )
Gdip_DisposeImage(pBitmapNeedle)
Gdip_UnlockBits(pBitmapHaystack,hBitmapData)
Return -1005
}
OuterX2 := ( !OuterX2 ? hWidth-nWidth+1 : OuterX2-nWidth+1 )
OuterY2 := ( !OuterY2 ? hHeight-nHeight+1 : OuterY2-nHeight+1 )
OutputCount := Gdip_MultiLockedBitsSearch(hStride,hScan,hWidth,hHeight
,nStride,nScan,nWidth,nHeight,OutputList,OuterX1,OuterY1,OuterX2,OuterY2
,Variation,SearchDirection,Instances,LineDelim,CoordDelim)
Gdip_UnlockBits(pBitmapHaystack,hBitmapData)
Gdip_UnlockBits(pBitmapNeedle,nBitmapData)
If ( DumpCurrentNeedle )
Gdip_DisposeImage(pBitmapNeedle)
Return OutputCount
}
Gdip_SetBitmapTransColor(pBitmap,TransColor) {
static _SetBmpTrans, Ptr, PtrA
if !( _SetBmpTrans ) {
Ptr := A_PtrSize ? "UPtr" : "UInt"
PtrA := Ptr . "*"
MCode_SetBmpTrans := "
            (LTrim Join
            8b44240c558b6c241cc745000000000085c07e77538b5c2410568b74242033c9578b7c2414894c24288da424000000
            0085db7e458bc18d1439b9020000008bff8a0c113a4e0275178a4c38013a4e01750e8a0a3a0e7508c644380300ff450083c0
            0483c204b9020000004b75d38b4c24288b44241c8b5c2418034c242048894c24288944241c75a85f5e5b33c05dc3,405
            34c8b5424388bda41c702000000004585c07e6448897c2410458bd84c8b4424304963f94c8d49010f1f800000000085db7e3
            8498bc1488bd3660f1f440000410fb648023848017519410fb6480138087510410fb6083848ff7507c640020041ff024883c
            00448ffca75d44c03cf49ffcb75bc488b7c241033c05bc3
)"
if ( A_PtrSize == 8 )
MCode_SetBmpTrans := SubStr(MCode_SetBmpTrans,InStr(MCode_SetBmpTrans,",")+1)
else
MCode_SetBmpTrans := SubStr(MCode_SetBmpTrans,1,InStr(MCode_SetBmpTrans,",")-1)
VarSetCapacity(_SetBmpTrans, LEN := StrLen(MCode_SetBmpTrans)//2, 0)
Loop, %LEN%
NumPut("0x" . SubStr(MCode_SetBmpTrans,(2*A_Index)-1,2), _SetBmpTrans, A_Index-1, "uchar")
MCode_SetBmpTrans := ""
DllCall("VirtualProtect", Ptr,&_SetBmpTrans, Ptr,VarSetCapacity(_SetBmpTrans), "uint",0x40, PtrA,0)
}
If !pBitmap
Return -2001
If TransColor not between 0 and 0xFFFFFF
Return -2002
Gdip_GetImageDimensions(pBitmap,W,H)
If !(W && H)
Return -2003
If Gdip_LockBits(pBitmap,0,0,W,H,Stride,Scan,BitmapData)
Return -2004
Gdip_FromARGB(TransColor,A,R,G,B), VarSetCapacity(TransColor,0), VarSetCapacity(TransColor,3,255)
NumPut(B,TransColor,0,"UChar"), NumPut(G,TransColor,1,"UChar"), NumPut(R,TransColor,2,"UChar")
MCount := 0
E := DllCall(&_SetBmpTrans, Ptr,Scan, "int",W, "int",H, "int",Stride, Ptr,&TransColor, "int*",MCount, "cdecl int")
Gdip_UnlockBits(pBitmap,BitmapData)
If ( E != 0 ) {
ErrorLevel := E
Return -2005
}
Return MCount
}
Gdip_MultiLockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride,nScan,nWidth,nHeight
,ByRef OutputList="",OuterX1=0,OuterY1=0,OuterX2=0,OuterY2=0,Variation=0
,SearchDirection=1,Instances=0,LineDelim="`n",CoordDelim=",")
{
OutputList := ""
OutputCount := !Instances
InnerX1 := OuterX1 , InnerY1 := OuterY1
InnerX2 := OuterX2 , InnerY2 := OuterY2
iX := 1, stepX := 1, iY := 1, stepY := 1
Modulo := Mod(SearchDirection,4)
If ( Modulo > 1 )
iY := 2, stepY := 0
If !Mod(Modulo,3)
iX := 2, stepX := 0
P := "Y", N := "X"
If ( SearchDirection > 4 )
P := "X", N := "Y"
iP := i%P%, iN := i%N%
While (!(OutputCount == Instances) && (0 == Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride
,nScan,nWidth,nHeight,FoundX,FoundY,OuterX1,OuterY1,OuterX2,OuterY2,Variation,SearchDirection)))
{
OutputCount++
OutputList .= LineDelim FoundX CoordDelim FoundY
Outer%P%%iP% := Found%P%+step%P%
Inner%N%%iN% := Found%N%+step%N%
Inner%P%1 := Found%P%
Inner%P%2 := Found%P%+1
While (!(OutputCount == Instances) && (0 == Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride
,nScan,nWidth,nHeight,FoundX,FoundY,InnerX1,InnerY1,InnerX2,InnerY2,Variation,SearchDirection)))
{
OutputCount++
OutputList .= LineDelim FoundX CoordDelim FoundY
Inner%N%%iN% := Found%N%+step%N%
}
}
OutputList := SubStr(OutputList,1+StrLen(LineDelim))
OutputCount -= !Instances
Return OutputCount
}
Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride,nScan,nWidth,nHeight
,ByRef x="",ByRef y="",sx1=0,sy1=0,sx2=0,sy2=0,Variation=0,sd=1)
{
static _ImageSearch, Ptr, PtrA
if !( _ImageSearch ) {
Ptr := A_PtrSize ? "UPtr" : "UInt"
PtrA := Ptr . "*"
MCode_ImageSearch := "
            (LTrim Join
            8b44243883ec205355565783f8010f857a0100008b7c2458897c24143b7c24600f8db50b00008b44244c8b5c245c8b
            4c24448b7424548be80fafef896c242490897424683bf30f8d0a0100008d64240033c033db8bf5896c241c895c2420894424
            183b4424480f8d0401000033c08944241085c90f8e9d0000008b5424688b7c24408beb8d34968b54246403df8d4900b80300
            0000803c18008b442410745e8b44243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424108b7c
            24408b4c24444083c50483c30483c604894424103bc17c818b5c24208b74241c0374244c8b44241840035c24508974241ce9
            2dffffff8b6c24688b5c245c8b4c244445896c24683beb8b6c24240f8c06ffffff8b44244c8b7c24148b7424544703e8897c
            2414896c24243b7c24600f8cd5feffffe96b0a00008b4424348b4c246889088b4424388b4c24145f5e5d890833c05b83c420
            c383f8020f85870100008b7c24604f897c24103b7c24580f8c310a00008b44244c8b5c245c8b4c24448bef0fafe8f7d88944
            24188b4424548b742418896c24288d4900894424683bc30f8d0a0100008d64240033c033db8bf5896c2420895c241c894424
            243b4424480f8d0401000033c08944241485c90f8e9d0000008b5424688b7c24408beb8d34968b54246403df8d4900b80300
            0000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424148b7c
            24408b4c24444083c50483c30483c604894424143bc17c818b5c241c8b7424200374244c8b44242440035c245089742420e9
            2dffffff8b6c24688b5c245c8b4c244445896c24683beb8b6c24280f8c06ffffff8b7c24108b4424548b7424184f03ee897c
            2410896c24283b7c24580f8dd5feffffe9db0800008b4424348b4c246889088b4424388b4c24105f5e5d890833c05b83c420
            c383f8030f85650100008b7c24604f897c24103b7c24580f8ca10800008b44244c8b6c245c8b5c24548b4c24448bf70faff0
            4df7d8896c242c897424188944241c8bff896c24683beb0f8c020100008d64240033c033db89742424895c2420894424283b
            4424480f8d76ffffff33c08944241485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d4900b80300
            0000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b4424148b7c
            24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e92bffffff
            8b6c24688b5c24548b4c24448b7424184d896c24683beb0f8d0affffff8b7c24108b44241c4f03f0897c2410897424183b7c
            24580f8c580700008b6c242ce9d4feffff83f8040f85670100008b7c2458897c24103b7c24600f8d340700008b44244c8b6c
            245c8b5c24548b4c24444d8bf00faff7896c242c8974241ceb098da424000000008bff896c24683beb0f8c020100008d6424
            0033c033db89742424895c2420894424283b4424480f8d06feffff33c08944241485c90f8e9f0000008b5424688b7c24408b
            eb8d34968b54246403dfeb038d4900b803000000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f
            752bca3bf97c6f8b44243c0fb64c06018b4424400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d
            04113bf87f3e2bca3bf97c388b4424148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b742424
            8b4424280374244c40035c2450e92bffffff8b6c24688b5c24548b4c24448b74241c4d896c24683beb0f8d0affffff8b4424
            4c8b7c24104703f0897c24108974241c3b7c24600f8de80500008b6c242ce9d4feffff83f8050f85890100008b7c2454897c
            24683b7c245c0f8dc40500008b5c24608b6c24588b44244c8b4c2444eb078da42400000000896c24103beb0f8d200100008b
            e80faf6c2458896c241c33c033db8bf5896c2424895c2420894424283b4424480f8d0d01000033c08944241485c90f8ea600
            00008b5424688b7c24408beb8d34968b54246403dfeb0a8da424000000008d4900b803000000803c03008b442414745e8b44
            243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b4424400fb67c28018d04113bf87f5a
            2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424148b7c24408b4c24444083c50483c30483
            c604894424143bc17c818b5c24208b7424240374244c8b44242840035c245089742424e924ffffff8b7c24108b6c241c8b44
            244c8b5c24608b4c24444703e8897c2410896c241c3bfb0f8cf3feffff8b7c24688b6c245847897c24683b7c245c0f8cc5fe
            ffffe96b0400008b4424348b4c24688b74241089088b4424385f89305e5d33c05b83c420c383f8060f85670100008b7c2454
            897c24683b7c245c0f8d320400008b6c24608b5c24588b44244c8b4c24444d896c24188bff896c24103beb0f8c1a0100008b
            f50faff0f7d88974241c8944242ceb038d490033c033db89742424895c2420894424283b4424480f8d06fbffff33c0894424
            1485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d4900b803000000803c03008b442414745e8b44
            243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c06018b4424400fb67c28018d04113bf87f56
            2bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b4424148b7c24408b4c24444083c50483c30483
            c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e92bffffff8b6c24108b74241c0374242c8b5c
            24588b4c24444d896c24108974241c3beb0f8d02ffffff8b44244c8b7c246847897c24683b7c245c0f8de60200008b6c2418
            e9c2feffff83f8070f85670100008b7c245c4f897c24683b7c24540f8cc10200008b6c24608b5c24588b44244c8b4c24444d
            896c241890896c24103beb0f8c1a0100008bf50faff0f7d88974241c8944242ceb038d490033c033db89742424895c242089
            4424283b4424480f8d96f9ffff33c08944241485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d49
            00b803000000803c18008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c
            06018b4424400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b44
            24148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e9
            2bffffff8b6c24108b74241c0374242c8b5c24588b4c24444d896c24108974241c3beb0f8d02ffffff8b44244c8b7c24684f
            897c24683b7c24540f8c760100008b6c2418e9c2feffff83f8080f85640100008b7c245c4f897c24683b7c24540f8c510100
            008b5c24608b6c24588b44244c8b4c24448d9b00000000896c24103beb0f8d200100008be80faf6c2458896c241c33c033db
            8bf5896c2424895c2420894424283b4424480f8d9dfcffff33c08944241485c90f8ea60000008b5424688b7c24408beb8d34
            968b54246403dfeb0a8da424000000008d4900b803000000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04
            113bf87f792bca3bf97c738b44243c0fb64c06018b4424400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0f
            b604068d0c103bf97f422bc23bf87c3c8b4424148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c2420
            8b7424240374244c8b44242840035c245089742424e924ffffff8b7c24108b6c241c8b44244c8b5c24608b4c24444703e889
            7c2410896c241c3bfb0f8cf3feffff8b7c24688b6c24584f897c24683b7c24540f8dc5feffff8b4424345fc700ffffffff8b
            4424345e5dc700ffffffffb85ff0ffff5b83c420c3,4c894c24204c89442418488954241048894c24085355565741544
            155415641574883ec188b8424c80000004d8bd94d8bd0488bda83f8010f85b3010000448b8c24a800000044890c24443b8c2
            4b80000000f8d66010000448bac24900000008b9424c0000000448b8424b00000008bbc2480000000448b9424a0000000418
            bcd410fafc9894c24040f1f84000000000044899424c8000000453bd00f8dfb000000468d2495000000000f1f80000000003
            3ed448bf933f6660f1f8400000000003bac24880000000f8d1701000033db85ff7e7e458bf4448bce442bf64503f7904d63c
            14d03c34180780300745a450fb65002438d040e4c63d84c035c2470410fb64b028d0411443bd07f572bca443bd17c50410fb
            64b01450fb650018d0411443bd07f3e2bca443bd17c37410fb60b450fb6108d0411443bd07f272bca443bd17c204c8b5c247
            8ffc34183c1043bdf7c8fffc54503fd03b42498000000e95effffff8b8424c8000000448b8424b00000008b4c24044c8b5c2
            478ffc04183c404898424c8000000413bc00f8c20ffffff448b0c24448b9424a000000041ffc14103cd44890c24894c24044
            43b8c24b80000000f8cd8feffff488b5c2468488b4c2460b85ff0ffffc701ffffffffc703ffffffff4883c418415f415e415
            d415c5f5e5d5bc38b8424c8000000e9860b000083f8020f858c010000448b8c24b800000041ffc944890c24443b8c24a8000
            0007cab448bac2490000000448b8424c00000008b9424b00000008bbc2480000000448b9424a0000000418bc9410fafcd418
            bc5894c2404f7d8894424080f1f400044899424c8000000443bd20f8d02010000468d2495000000000f1f80000000004533f
            6448bf933f60f1f840000000000443bb424880000000f8d56ffffff33db85ff0f8e81000000418bec448bd62bee4103ef496
            3d24903d3807a03007460440fb64a02418d042a4c63d84c035c2470410fb64b02428d0401443bc87f5d412bc8443bc97c554
            10fb64b01440fb64a01428d0401443bc87f42412bc8443bc97c3a410fb60b440fb60a428d0401443bc87f29412bc8443bc97
            c214c8b5c2478ffc34183c2043bdf7c8a41ffc64503fd03b42498000000e955ffffff8b8424c80000008b9424b00000008b4
            c24044c8b5c2478ffc04183c404898424c80000003bc20f8c19ffffff448b0c24448b9424a0000000034c240841ffc9894c2
            40444890c24443b8c24a80000000f8dd0feffffe933feffff83f8030f85c4010000448b8c24b800000041ffc944898c24c80
            00000443b8c24a80000000f8c0efeffff8b842490000000448b9c24b0000000448b8424c00000008bbc248000000041ffcb4
            18bc98bd044895c24080fafc8f7da890c24895424048b9424a0000000448b542404458beb443bda0f8c13010000468d249d0
            000000066660f1f84000000000033ed448bf933f6660f1f8400000000003bac24880000000f8d0801000033db85ff0f8e960
            00000488b4c2478458bf4448bd6442bf64503f70f1f8400000000004963d24803d1807a03007460440fb64a02438d04164c6
            3d84c035c2470410fb64b02428d0401443bc87f63412bc8443bc97c5b410fb64b01440fb64a01428d0401443bc87f48412bc
            8443bc97c40410fb60b440fb60a428d0401443bc87f2f412bc8443bc97c27488b4c2478ffc34183c2043bdf7c8a8b8424900
            00000ffc54403f803b42498000000e942ffffff8b9424a00000008b8424900000008b0c2441ffcd4183ec04443bea0f8d11f
            fffff448b8c24c8000000448b542404448b5c240841ffc94103ca44898c24c8000000890c24443b8c24a80000000f8dc2fef
            fffe983fcffff488b4c24608b8424c8000000448929488b4c2468890133c0e981fcffff83f8040f857f010000448b8c24a80
            0000044890c24443b8c24b80000000f8d48fcffff448bac2490000000448b9424b00000008b9424c0000000448b8424a0000
            0008bbc248000000041ffca418bcd4489542408410fafc9894c2404669044899424c8000000453bd00f8cf8000000468d249
            5000000000f1f800000000033ed448bf933f6660f1f8400000000003bac24880000000f8df7fbffff33db85ff7e7e458bf44
            48bce442bf64503f7904d63c14d03c34180780300745a450fb65002438d040e4c63d84c035c2470410fb64b028d0411443bd
            07f572bca443bd17c50410fb64b01450fb650018d0411443bd07f3e2bca443bd17c37410fb60b450fb6108d0411443bd07f2
            72bca443bd17c204c8b5c2478ffc34183c1043bdf7c8fffc54503fd03b42498000000e95effffff8b8424c8000000448b842
            4a00000008b4c24044c8b5c2478ffc84183ec04898424c8000000413bc00f8d20ffffff448b0c24448b54240841ffc14103c
            d44890c24894c2404443b8c24b80000000f8cdbfeffffe9defaffff83f8050f85ab010000448b8424a000000044890424443
            b8424b00000000f8dc0faffff8b9424c0000000448bac2498000000448ba424900000008bbc2480000000448b8c24a800000
            0428d0c8500000000898c24c800000044894c2404443b8c24b80000000f8d09010000418bc4410fafc18944240833ed448bf
            833f6660f1f8400000000003bac24880000000f8d0501000033db85ff0f8e87000000448bf1448bce442bf64503f74d63c14
            d03c34180780300745d438d040e4c63d84d03da450fb65002410fb64b028d0411443bd07f5f2bca443bd17c58410fb64b014
            50fb650018d0411443bd07f462bca443bd17c3f410fb60b450fb6108d0411443bd07f2f2bca443bd17c284c8b5c24784c8b5
            42470ffc34183c1043bdf7c8c8b8c24c8000000ffc54503fc4103f5e955ffffff448b4424048b4424088b8c24c80000004c8
            b5c24784c8b54247041ffc04103c4448944240489442408443b8424b80000000f8c0effffff448b0424448b8c24a80000004
            1ffc083c10444890424898c24c8000000443b8424b00000000f8cc5feffffe946f9ffff488b4c24608b042489018b4424044
            88b4c2468890133c0e945f9ffff83f8060f85aa010000448b8c24a000000044894c2404443b8c24b00000000f8d0bf9ffff8
            b8424b8000000448b8424c0000000448ba424900000008bbc2480000000428d0c8d00000000ffc88944240c898c24c800000
            06666660f1f840000000000448be83b8424a80000000f8c02010000410fafc4418bd4f7da891424894424084533f6448bf83
            3f60f1f840000000000443bb424880000000f8df900000033db85ff0f8e870000008be9448bd62bee4103ef4963d24903d38
            07a03007460440fb64a02418d042a4c63d84c035c2470410fb64b02428d0401443bc87f64412bc8443bc97c5c410fb64b014
            40fb64a01428d0401443bc87f49412bc8443bc97c41410fb60b440fb60a428d0401443bc87f30412bc8443bc97c284c8b5c2
            478ffc34183c2043bdf7c8a8b8c24c800000041ffc64503fc03b42498000000e94fffffff8b4424088b8c24c80000004c8b5
            c247803042441ffcd89442408443bac24a80000000f8d17ffffff448b4c24048b44240c41ffc183c10444894c2404898c24c
            8000000443b8c24b00000000f8ccefeffffe991f7ffff488b4c24608b4424048901488b4c246833c0448929e992f7ffff83f
            8070f858d010000448b8c24b000000041ffc944894c2404443b8c24a00000000f8c55f7ffff8b8424b8000000448b8424c00
            00000448ba424900000008bbc2480000000428d0c8d00000000ffc8890424898c24c8000000660f1f440000448be83b8424a
            80000000f8c02010000410fafc4418bd4f7da8954240c8944240833ed448bf833f60f1f8400000000003bac24880000000f8
            d4affffff33db85ff0f8e89000000448bf1448bd6442bf64503f74963d24903d3807a03007460440fb64a02438d04164c63d
            84c035c2470410fb64b02428d0401443bc87f63412bc8443bc97c5b410fb64b01440fb64a01428d0401443bc87f48412bc84
            43bc97c40410fb60b440fb60a428d0401443bc87f2f412bc8443bc97c274c8b5c2478ffc34183c2043bdf7c8a8b8c24c8000
            000ffc54503fc03b42498000000e94fffffff8b4424088b8c24c80000004c8b5c24780344240c41ffcd89442408443bac24a
            80000000f8d17ffffff448b4c24048b042441ffc983e90444894c2404898c24c8000000443b8c24a00000000f8dcefeffffe
            9e1f5ffff83f8080f85ddf5ffff448b8424b000000041ffc84489442404443b8424a00000000f8cbff5ffff8b9424c000000
            0448bac2498000000448ba424900000008bbc2480000000448b8c24a8000000428d0c8500000000898c24c800000044890c2
            4443b8c24b80000000f8d08010000418bc4410fafc18944240833ed448bf833f6660f1f8400000000003bac24880000000f8
            d0501000033db85ff0f8e87000000448bf1448bce442bf64503f74d63c14d03c34180780300745d438d040e4c63d84d03da4
            50fb65002410fb64b028d0411443bd07f5f2bca443bd17c58410fb64b01450fb650018d0411443bd07f462bca443bd17c3f4
            10fb603450fb6108d0c10443bd17f2f2bc2443bd07c284c8b5c24784c8b542470ffc34183c1043bdf7c8c8b8c24c8000000f
            fc54503fc4103f5e955ffffff448b04248b4424088b8c24c80000004c8b5c24784c8b54247041ffc04103c44489042489442
            408443b8424b80000000f8c10ffffff448b442404448b8c24a800000041ffc883e9044489442404898c24c8000000443b842
            4a00000000f8dc6feffffe946f4ffff8b442404488b4c246089018b0424488b4c2468890133c0e945f4ffff
)"
if ( A_PtrSize == 8 )
MCode_ImageSearch := SubStr(MCode_ImageSearch,InStr(MCode_ImageSearch,",")+1)
else
MCode_ImageSearch := SubStr(MCode_ImageSearch,1,InStr(MCode_ImageSearch,",")-1)
VarSetCapacity(_ImageSearch, LEN := StrLen(MCode_ImageSearch)//2, 0)
Loop, %LEN%
NumPut("0x" . SubStr(MCode_ImageSearch,(2*A_Index)-1,2), _ImageSearch, A_Index-1, "uchar")
MCode_ImageSearch := ""
DllCall("VirtualProtect", Ptr,&_ImageSearch, Ptr,VarSetCapacity(_ImageSearch), "uint",0x40, PtrA,0)
}
If ( sx2 < sx1 )
return -3001
If ( sy2 < sy1 )
return -3002
If ( sx2 > (hWidth-nWidth+1) )
return -3003
If ( sy2 > (hHeight-nHeight+1) )
return -3004
If ( sx2-sx1 == 0 )
return -3005
If ( sy2-sy1 == 0 )
return -3006
x := 0, y := 0
, E := DllCall( &_ImageSearch, "int*",x, "int*",y, Ptr,hScan, Ptr,nScan, "int",nWidth, "int",nHeight
, "int",hStride, "int",nStride, "int",sx1, "int",sy1, "int",sx2, "int",sy2, "int",Variation
, "int",sd, "cdecl int")
Return ( E == "" ? -3007 : E )
}
#InstallKeybdHook
#InstallMouseHook
ListLines Off
Version_Number = 3.43
Global Serialz, Clippy, Title, sleepstill, currMon
Checkp=0
Sleepstill = 0
LoopCount = 0
ButtonLoop = 0
Searchcount = 0
Guicount = 1
Serialcount = 0
Programend = 0
Needrefresh = 0
Editfield2 =
Complete = 0
Modifier =
ToPause = 0
badserial = 0
Serialzcounter2 = 0
Serialzcounter = 0
Prefixcount = 5
TotalPrefixes = -2
Textaddbutton = Please Shift + mouse button click on the "Add Button" in the ACM effectivity screen to get its position.
Textprefixbutton = Please Shift + mouse button click in the "prefix" edit field in the ACM effectivity screen to get it's location.
Textapplybutton = Please Shift + mouse button click on the "Apply button" in the ACM effectivity screen to get it's location.
Radiobutton = 1
IfNotExist, C:\SerialMacro
{
FileCreateDir, C:\SerialMacro
sleep 100
}
IfnotExist, C:\Serialmacro\Settings.ini
{
FileAppend,  C:\Serialmacro\Settings.ini
sleep 1000
IniWrite, 20,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
}
IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
FileInstall, C:\ArbortextMacrosn\serial macro\How to use Serial Macros.pdf, C:\SerialMacro\How to use Serial Macros.pdf,1
IfNotExist, C:\SerialMacro\redimage.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\redimage.png, C:\SerialMacro\redimage.png,1
}
IfNotExist, C:\SerialMacro\plussign.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\plussign.png, C:\SerialMacro\plussign.png,1
}
IfNotExist, C:\SerialMacro\Activeplus.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\activeplus.png, C:\SerialMacro\activeplus.png,1
}
IfNotExist, C:\SerialMacro\orange button.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\orange button.png, C:\SerialMacro\orange button.png,1
}
IfNotExist, C:\SerialMacro\paused.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\paused.png, C:\SerialMacro\paused.png,1
}
IfNotExist, C:\SerialMacro\start.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\start.png, C:\SerialMacro\start.png,1
}
IfNotExist, C:\SerialMacro\Running.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\Running.png, C:\SerialMacro\Running.png,1
}
IfNotExist, C:\SerialMacro\Stopped.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\Stopped.png, C:\SerialMacro\Stopped.png,1
}
IfNotExist, C:\SerialMacro\background.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\background.png, C:\SerialMacro\background.png,1
}
IfnotExist C:\SerialMacro\uredimage.png
{
FileNamered := "C:\SerialMacro\redimage.png"
}
Else
{
FileNamered := "C:\SerialMacro\uredimage.png"
}
IfnotExist C:\SerialMacro\uplussign.png
{
FileNameSearch := "C:\SerialMacro\plussign.png"
}
Else
{
FileNameSearch := "C:\SerialMacro\uplussign.png"
}
IfnotExist C:\SerialMacro\uActiveplus.png
{
FileNameCheck := "C:\SerialMacro\Activeplus.png"
}
Else
{
FileNameCheck := "C:\SerialMacro\uActiveplus.png"
}
IfNotExist C:\SerialMacro\uorange button.png
{
FileNameButton := "C:\SerialMacro\orange button.png"
}
Else
{
FileNameButton := "C:\SerialMacro\uorange button.png"
}
Menu, Tray, NoStandard
Menu, Tray, Add, How to use, HowTo
Menu Tray, Add, Check For update, Versioncheck
Menu, Tray, Add, Quit, Quitapp
gosub, Updatechecker
gosub, SerialsGUIscreen
pToken := Gdip_Startup()
bmpNeedle1 := Gdip_CreateBitmapFromFile(FileNamered)
bmpNeedleSearch := Gdip_CreateBitmapFromFile(FileNameSearch)
bmpNeedleCheck := Gdip_CreateBitmapFromFile(FileNameCheck)
bmpNeedleB := Gdip_CreateBitmapFromFile(FileNameButton)
$^2::
{
Gosub, Enterallserials
skipbox = 0
}
^q::
{
Gosub, Exitprogram
Return
}
Enterallserials:
{
Refreshrate := 20
Prefixcount = 5
addtime = 0
Badlist =
StartTime := A_TickCount
msgbox,262144,,Please maximize the ACM window to help prevent problems.`n`n Press the OK button once that is complete.
Prefixcount = 5
Badlist =
Guicontrol,hide, Starting
Gui, Submit, NoHide
Stoptimer = 0
Serialzcounter2 = 0
Serialzcounter = 0
Gosub, Getmousepositions
If Stop = 1
{
stop = 0
gosub, Getmousepositions
}
If Stop = 1
{
stop = 0
gosub, Getmousepositions
}
If Stop = 1
{
stop = 0
gosub, Getmousepositions
}
If Stop = 1
{
stop = 0
gosub, Getmousepositions
}
If Stop = 1
{
stop = 0
gosub, Getmousepositions
}
Gosub, Comma_Check
Gosub, Searchend
Loop
{
Gosub, checkforactivity
If breakloop = 1
{
Gui 1: -AlwaysOnTop
Break
SplashTextOn,,,Macro Stopped
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
}
Gosub, loopcounts
If LoopCount >= %Refreshrate%
{
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
Sleep 1000
Needrefresh = 0
sleep 300
Gosub, Searchend
Gosub, Win_check
sleep 100
Send {F5}
Sleep 2000
Searchcount = 0
Searchcountser = 0
Gosub Refreshpage
Loopcount = 0
Searchcount = 0
Searchcountser = 0
Needrefresh = 0
Sleep 1000
}
Sleep 100
Serialz=0
Searchcount = 0
Searchcountser = 0
Gosub, Get_Prefix
Gosub, Copy_Serial
Gosub, Number_Check
}
return
}
Getmousepositions:
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
SetTimer, ToolTipTimerbutton, 10
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Keywait, Shift,D
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Gosub, KeystateTimer
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
MouseGetPos, Searchx, Searchy
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
WinGetTitle, Title, A
currMon := GetCurrentMonitor()
Sleep 500
tooltip,
SetTimer,ToolTipTimerbutton,off
Sleep 1000
settimer, ToolTipTimerprefix,10
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Keywait, Shift,D,t5
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Gosub, KeystateTimer
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
MouseGetPos, prefixx, prefixy
tooltip,
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
settimer, ToolTipTimerprefix,Off
Sleep 1000
SetTimer, ToolTipTimerapply, 10
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Keywait, Shift,D,t5
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
Gosub, KeystateTimer
If stop = 1
{
SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,
return
}
MouseGetPos, Applyx, Applyy
tooltip,
SetTimer,ToolTipTimerapply,off
Sleep 500
ToolTip,
Click, %Searchx%, %Searchy%
Sleep 500
Click, %prefixx%, %prefixy%
Sleep 1000
mousemove, 300,300
Return
}
Get_Prefix:
{
Modifier =
Guicount++
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Enter Serial Macro
sleep 50
COntrolSEnd,,{Shift Down}{Right 3}{Shift Up}, Enter Serial Macro
ControlGet,Prefixes,Selected,,,Enter Serial Macro
sleep 300
If Prefixes = ***
{
complete = 1
}
COntrolSEnd,, {del}, Enter Serial Macro
Text1 = %Prefixes%
Guitextlocation = %Prefixes%
if Prefixcount != 5
{
PrefixStore1 = %PrefixStore%
Prefixstore =  %Guitextlocation%
}
If Prefixcount = 5
{
PrefixStore1 = %Guitextlocation%
Prefixstore = %Guitextlocation%
Prefixcount = 1
}
Return
}
Copy_Serial:
{
Serialz++
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Enter Serial Macro
sleep 50
COntrolSEnd,,{Shift Down}{Right 5}{Shift Up},Enter Serial Macro
ControlGet, SerialN, Selected,,,Enter Serial Macro
COntrolSEnd,, {del}, Enter Serial Macro
If serialz = 1
{
Clippy = %SerialN%
If Serialzcounter  = 0
{
If Serialstore1 =
{
Serialstore1 = %Clippy%
}
Serialstore = %Clippy%
Serialzcounter = 1
}
Else
{
Serialstore1 = %Serialstore%
Serialstore =  %Clippy%
}
}
If Serialz = 2
{
Clippy2 = %SerialN%
If Serialzcounter2  = 0
{
If Serialstore3 =
{
Serialstore3 = %Clippy2%
}
Serialstore2 = %Clippy2%
Serialzcounter2 = 1
}
Else
{
Serialstore3 = %Serialstore2%
Serialstore2 =  %Clippy2%
}
COntrolSEnd,, {del 2}, Enter Serial Macro
GOsub, Upcheck
}
Return
}
loopcounts:
{
LoopCount++
return
}
Number_Check:
{
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Enter Serial Macro
sleep 50
COntrolsend,,{Shift down}{RIght}{SHift up},Enter Serial Macro
ControlGet,Numbers,Selected,,,Enter Serial Macro
COntrolSEnd,, {del},Enter Serial Macro
Text3 = %Numbers%
If Text3 = -
{
Gosub, Copy_Serial
If Clippy2 =
{
Clippytemp = 99999
}
else {
clippytemp = %clippy2%
}
Searchcount = 0
Searchcountser = 0
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippytemp%
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
Gui, submit, nohide
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}
Else If Text3 = ,
{
If Serialz = 1
{
Clippy2 = %Clippy%
Searchcount = 0
Searchcountser = 0
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippy2%
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Enter Serial Macro
Gui, submit, nohide
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}
Else if Serialz = 2
{
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippy2%
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Enter Serial Macro
Gui, submit, nohide
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}
}
Else if Text3 =
{
If Serialz = 1
{
Clippy2 = %Clippy%
If Clippy2 =
{
Clippytemp = 99999
}
else {
clippytemp = %clippy2%
}
Searchcount = 0
Searchcountser = 0
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippytemp%
if complete = 1
GuiControl,1:,nextserialtoadd,
Else
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Enter Serial Macro
Gui, submit, nohide
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippytemp%
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Enter Serial Macro
Gui, submit, nohide
Gosub, Searchendforserials
Gosub, EnterSerials
Return
}
return
}
Comma_Check:
{
COntrolSEnd,, {Shift Down}{Right}{SHift Up},Enter Serial Macro
ControlGet,Commacheck,Selected,,,Enter Serial Macro
If Commacheck = ,
{
Controlsend,,{BackSpace 2},Enter Serial Macro
Return
}
Else
{
COntrolSEnd,, {Left}{BackSpace 2}, Enter Serial Macro
Return
}
Return
}
Check_Space:
{
If Clippy =
{
ControlSEnd,,{Del},Enter Serial Macro
ExitApp
}
Return
}
Reloading:
{
Send {Shift Up}{Ctrl Up}
Msgbox,262144,Reload -- Enter Serial Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Serial Macro is currently Paused.`n`n Press Okay Button to reload Macro.`n`n Be sure to View the Serials Already Added Section to view the  Serial Numbers that were added into ACM to figure out where to start over if the macro messed up.
Pause, on
Reload
Return
}
#IfWinActive ahk_class TTAFrameXClass
Esc::
{
Gosub,Exitprogram
Return
}
#ifwinactive Enter Serial Macro V
Esc::
{
Gosub,Exitprogram
Return
}
#ifwinactive
Exitprogram:
{
Msgbox,262148,Stop -- Enter Serial Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.
ifmsgbox yes
{
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Send {Shift Up}{Ctrl Up}
breakloop = 1
Exit
}
Else
Return
}
UPcheck:
{
If Clippy2 = 99999
Clippy2 = %A_Space%
Return
}
Quitapp:
{
msgbox,262148,Quit -- Enter Serial Macro V%Version_Number%, Are you sure you want to quit?
ifMsgBox Yes
{
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Send {Shift Up}{Ctrl Up}
breakloop = 1
WinSet, AlwaysOnTop, Off,%Title%
ExitApp
}
Else
{
Return
}
Return
}
Win_check:
{
IfWinNotActive , %Title%
{
WinActivate, %Title%
WinWaitActive, %Title%,,3
Sleep 500
}
return
}
Searchend:
{
listlines off
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100
If RETSearch < 0
{
if RETSearch = -1001
RETSearch = invalid haystack or needle bitmap pointer
if RETSearch = -1002
RETSearch = invalid variation value
if RETSearch = -1003
RETSearch = Unable to lock haystack bitmap bits
if RETSearch = -1004
RETSearch = Unable to lock needle bitmap bits
if RETSearch = -1005
RETSearch = Cannot find monitor for screen capture
Msgbox,262144, Enter Serial Macro V%Version_Number%, Error Searchend (bmpNeedle1) %RetSearch%
Exit
}
If RETSearch > 0
{
SetTimer, refreshcheck, Off
Refreshchecks = 0
SleepStill = 0
}
If RETSearch = 0
{
If Searchcount = 7
{
Timeout = Yes
gosub, Macrotimedout
}
Sleepstill = 1
Sleep 500
searchcount++
Gosub, Searchend
}
Return
}
Searchendforserials:
{
listlines off
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100
If RETSearch < 0
{
if RETSearch = -1001
RETSearch = invalid haystack or needle bitmap pointer
if RETSearch = -1002
RETSearch = invalid variation value
if RETSearch = -1003
RETSearch = Unable to lock haystack bitmap bits
if RETSearch = -1004
RETSearch = Unable to lick needle bitmap bits
if RETSearch = -1005
RETSearch = Cannot find monitor for screen capture
Msgbox,262144, Enter Serial Macro V%Version_Number%, Error Searchend (bmpNeedle1) %RetSearch%
Exit
}
If RETSearch > 0
{
SplashTextOff
Return
}
If RETSearch = 0
{
listlines off
If addtime = 2
{
If Searchcountser = 1
{
SplashTextOn,,25,Serial Macro, Time out in 7
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 2
{
SplashTextOn,,25,Serial Macro, Time out in 6
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 3
{
SplashTextOn,,25,Serial Macro, Time out in 5
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 4
{
SplashTextOn,,25,Serial Macro, Time out in 4
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 5
{
SplashTextOn,,25,Serial Macro, Time out in 3
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 6
{
SplashTextOn,,25,Serial Macro, Time out in 2
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 7
{
SplashTextOn,,25,Serial Macro, Time out in 1
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 8
{
SplashTextoff
Timeout = Yes
gosub, Macrotimedout
}
}
If addtime = 1
{
If Searchcountser = 1
{
SplashTextOn,,25,Serial Macro, Time out in 6
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 2
{
SplashTextOn,,25,Serial Macro, Time out in 5
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 3
{
SplashTextOn,,25,Serial Macro, Time out in 4
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 4
{
SplashTextOn,,25,Serial Macro, Time out in 3
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 5
{
SplashTextOn,,25,Serial Macro, Time out in 2
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 6
{
SplashTextOn,,25,Serial Macro, Time out in 1
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 7
{
SplashTextoff
Timeout = Yes
gosub, Macrotimedout
}
}
Else
{
If Searchcountser = 1
{
SplashTextOn,,25,Serial Macro, Time out in 5
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 2
{
SplashTextOn,,25,Serial Macro, Time out in 4
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 3
{
SplashTextOn,,25,Serial Macro, Time out in 3
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 4
{
SplashTextOn,,25,Serial Macro, Time out in 2
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 5
{
SplashTextOn,,25,Serial Macro, Time out in 1
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 6
{
SplashTextoff
Timeout = Yes
gosub, Macrotimedout
}
}
Sleepstill = 1
Sleep 500
searchcountser++
Gosub, Searchendforserials
}
Return
}
refreshcheck:
{
Refreshchecks++
If refreshchecks = 10
{
Needrefresh = 1
Refreshchecks = 0
}
Return
}
GetCurrentMonitor()
{
SysGet, numberOfMonitors, MonitorCount
WinGetPos, winX, winY, winWidth, winHeight, A
winMidX := winX + winWidth / 2
winMidY := winY + winHeight / 2
Loop %numberOfMonitors%
{
SysGet, monArea, Monitor, %A_Index%
if (winMidX > monAreaLeft && winMidX < monAreaRight && winMidY < monAreaBottom && winMidY > monAreaTop)
return A_Index
}
SysGet, MonitorPrimary, MonitorPrimary
return "No Monitor Found"
}
Refreshpage:
{
Sleep 500
Click, %Searchx%, %Searchy%
Sleep 2000
Gosub, SearchendRefresh
Return
}
SearchendRefresh:
{
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100
If RETSearch < 0
{
Msgbox,262144,Enter Serial Macro V%Version_Number%, Error SearchendRefresh (bmpNeedle1) %RetSearch%
Exit
}
If RETSearch > 0
{
SleepStill = 0
Return
}
If RETSearch = 0
{
searchcount++
Gosub, Refreshpage
}
Return
}
Screen_check:
{
listlines off
bmpHaystack1 := Gdip_BitmapFromScreen(currMon)
ScreenSearch := Gdip_ImageSearch(bmpHaystack1,bmpNeedleSearch,,0,0,0,0,5,0,0,0)
If ScreenSearch < 0
{
Msgbox,262144, Enter Serial Macro V%Version_Number%, Error screensearch (bmpNeedleSearch) is  %ScreenSearch%
Reload
}
If ScreenSearch > 0
{
Return
}
If ScreenSearch = 0
{
ScreenCheck := Gdip_ImageSearch(bmpHaystack1,bmpNeedleCheck,,0,0,0,0,5,0,0,0)
If ScreenCheck < 0
{
Msgbox,262144, Enter Serial Macro V%Version_Number%, Error  screen check  (bmpNeedleCheck) is %ScreenCheck%
Reload
}
If ScreenCheck > 0
{
Return
}
If ScreenCheck = 0
{
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Msgbox,262144,Enter Serial Macro V%Version_Number%, OOPS!!`n`n Something bad happened and Cannot find the Effectivity screen.  `n`nMacro is now paused. `n`nMake the effectivity screen active and what is most likely something over the Effectivity screen. Hide that panel and find where I left off. `n`n Press Pause to unpause script after you press OK on this window.
Pause, on
}
Return
}
Return
}
Enterserials:
{
Gosub, Win_check
Loop, parse, Badlist, `,,all
{
If prefixes = %A_LoopField%
{
noacmPrefix = %prefixes%
Skipserial = 1
Break
}
else
Skipserial = 0
}
Loop, parse, DualENG, `,,all
{
If prefixes = %A_LoopField%
{
DUalACMCheck = 1
DualACMPrefix = %prefixes%
Break
}
Else
{
DUalACMCheck = 0
}
}
if Skipserial = 1
{
Modifier = --- *Serial Not in ACM*
Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%Clippy%-%Clippy2%%Modifier%`n
Guicontrol,hide, Editfield2,
GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%
Modifier =
Skipserial = 0
return
}
If Complete = 1
{
ElapsedTime := A_TickCount - StartTime
res1 := milli2hms(ElapsedTime, h, m, s)
Sleep 500
Send {f5}
Msgbox,262144,Enter Serial Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Macro Finished due to no more Serials to add. `n`n It took the macro %res1% to perform tasks. `n`n Please close Serial Macro Window when finished checking to ensure serials were entered correctly.
Guicontrol,1:, Editfield,
gosub, radio2h
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Exit
Return
}
Gosub, Screen_Check
Click, %prefixx%, %prefixy%
sleep 100
mousemove 300,300
sleep 100
SEndRaw, %text1%
Sleep 100
Send {Tab}
Gosub, Win_check
Sendraw, %Clippy%
sleep 250
Send {Tab}
Gosub, Win_check
SendRaw, %Clippy2%
Sleep 250
Send {Tab}
Sleep 650
If DUalACMCheck = 1
{
Modifier = --- **Multiple Engineering Models**
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, This Serial has multiple Engineering Models. Please select the model in ACM and press the OK button in this window.
SplashTextOn,,25,Serial Macro, Macro will resume in 2
sleep 1000
SplashTextOn,,25,Serial Macro, Macro will resume in 1
sleep 1000
SplashTextOff
gosub, Win_check
Click, %prefixx%, %prefixy%
sleep 500
Send {Tab 3}
sleep 500
DualACMCheck = 0
SleepStill = 0
sleep 250
}
Send {enter 2}
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
Serialcount +=1
Sleep 100
Searchcount = 0
Searchcountser = 0
If Clippy2 =
{
Clippy2 = 99999
}
Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%Clippy%-%Clippy2%%Modifier%`n
Guicontrol,hide, Editfield2,
GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%
Refreshchecks = 0
SleepStill = 0
Modifier =
Skipserial = 0
Return
}
milli2hms(milli, ByRef hours=0, ByRef mins=0, ByRef secs=0, secPercision=0)
{
SetFormat, FLOAT, 0.%secPercision%
milli /= 1000.0
secs := mod(milli, 60)
secs = %secs%Sec
SetFormat, FLOAT, 0.0
milli //= 60
mins := mod(milli, 60)
mins = %mins%Min
hours := milli //60
Hours = %hours%hrs
return hours . ":" . mins . ":" . secs
}
EXIT_LABEL:
{
If pToken !=0
{
Gdip_Shutdown(pToken)
}
ExitApp
}
throwout:
{
badserial = 1
Return
}
ToolTipTimerbutton:
{
tooltip, %Textaddbutton%
Return
}
ToolTipTimerapply:
{
tooltip, %Textapplybutton%
Return
}
ToolTipTimerprefix:
{
tooltip, %Textprefixbutton%
Return
}
KeystateTimer:
{
GetKeyState,Rmousee, Rbutton
GetKeyState,Lmousee, Lbutton
If Lmousee = D
{
Return
}
IF Rmousee = D
{
Return
}
Else
{
Sleep 100
GOsub, KeystateTimer
}
Return
}
Engmodel:
{
warned = 1
Gui, 3:Destroy
Modifier = --- **Multiple Engineering Models**
DUalACMCheck = 1
DualENG = %DualENG%%PrefixStore1%`,
StringTrimRight, EditField2, Editfield2,1
Editfield2n = %Editfield2%%Modifier%`n
EditField2 = %EditField2n%
Guicontrol, 1:,Editfield2, %Editfield2%
Gui 1: +alwaysontop
Msgbox,262144, Enter Serial Macro V%Version_Number%, Select the engeering model and then press the OK button on this window. `n`n The prefix will be logged for this session to prevent more timouts of the same prefix.
sleep 100
gosub, Win_check
Click, %prefixx%, %prefixy%
Send {Tab 3}
SplashTextOn,,25,Serial Macro, Macro will resume in 3
sleep 250
SplashTextOn,,25,Serial Macro, Macro will resume in 2
sleep 250
SplashTextOn,,25,Serial Macro, Macro will resume in 1
sleep 250
SplashTextOff
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, Off
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
modifier =
gosub, searchend
Return
}
Anarchy:
{
Gui, 3:Destroy
Msgbox,262148,-- Enter Serial Macro V%Version_Number%, Press Yes button if you found the issue and to Resume the Macro, No button to reload macro
IfMsgBox Yes
{
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, Off
}
Else
{
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Pause, Off
Gosub, Reloading
}
return
}
Serialnogo:
{
Gui, 3:Destroy
Modifier := --- *Serial Not in ACM*
Serialcount--
Badlist = %badlist%%PrefixStore1%`,
SplashTextOn,,25,Serial Macro, Macro will resume in 3
sleep 500
SplashTextOn,,25,Serial Macro, Macro will resume in 2
sleep 500
SplashTextOn,,25,Serial Macro, Macro will resume in 1
sleep 500
SplashTextOff
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, 1:Submit, NoHide
Pause, Off
gosub, Win_check
Click, %prefixx%, %prefixy%
sleep 500
Send {ctrl down}{a}{Ctrl up}
sleep 100
Send {Del}{Tab}
Send {Del}{Tab}
Send {Del}{Tab}
Click, %prefixx%, %prefixy%
Send {Del}
Modifier = --- *Serial Not in ACM*
StringTrimRight, EditField2, Editfield2,1
Editfield2n = %Editfield2%%Modifier%`n
EditField2 = %EditField2n%
Guicontrol, 1:,Editfield2, %Editfield2%
Gui 1: +alwaysontop
Modifier =
Return
}
pausedstate:
{
Sleep 2000
return
}
checkforactivity:
{
if A_TimeIdlePhysical < 4999
{
Gui 1: -AlwaysOnTop
{
If (A_TimeIdlePhysical > 0) and (A_TimeIdlePhysical < 1000)
Timeleft = 5
If (A_TimeIdlePhysical > 1000) and (A_TimeIdlePhysical < 2000)
Timeleft = 4
If (A_TimeIdlePhysical > 2000) and (A_TimeIdlePhysical < 3000)
Timeleft = 3
If (A_TimeIdlePhysical > 3000) and (A_TimeIdlePhysical < 4000)
Timeleft = 2
If (A_TimeIdlePhysical > 4000) and (A_TimeIdlePhysical < 5000)
Timeleft = 1
}
SplashTextOn,350,50,Macro paused, Macro is now paused due to user activity.`n Macro will resume after %timeleft% seconds of no user input
pausedstate = yes
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Sleep 1000
Gosub, checkforactivity
}
if A_TimeIdlePhysical > 5000
{
If pausedstate = yes
{
splashtextoff
Gui 1: +AlwaysOnTop
Gosub, Win_check
pausedstate = no
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
gosub, radio1h
Gui, Submit, NoHide
gosub, SerialFullScreen
sleep 1000
}
return
}
return
}
Insert::
Pause::
{
if A_IsPaused = 0
{
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,10
Pause, toggle, 1
Return
}
Else
{
gosub, radio1h
Gui 1: +AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, toggle, 1
}
return
}
pausesub:
{
if A_IsPaused = 0
{
gosub, radio1h
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Msgbox,262144,-- Enter Serial Macro V%Version_Number%, Macro is paused. Press pause to unpause,10
Pause, toggle, 1
Return
}
Else
{
Gui 1: +AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, toggle, 1
}
return
}
SerialFullScreen:
{
WinGetPos, Xarbor,yarbor,warbor,harbor, %title%
CurrmonAM := GetCurrentMonitor()
SysGet,Aarea,MonitorWorkArea,%currmonAM%
WidthA := AareaRight- AareaLeft
HeightA := aareaBottom - aAreaTop
leftt := aAreaLeft - 4
topp := AAreaTop - 4
MouseGetPos mmx,mmy
If yarbor = %topp%
{
If xarbor = %leftt%
Return
}
Else
CoordMode, mouse, Relative
MouseMove 300,10
Click
Click
Coordmode, mouse, screen
return
}
acmlong:
{
pause, off
Gui 1: +AlwaysOnTop
gosub, radio1h
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui 1:Submit, NoHide
Gui, 3:Destroy
if addtime = 0
{
Addtimetemp++
If addtimetemp = 2
{
Addtime = 1
Addtimetemp = 0
Return
}
}
If addtime = 1
{
Addtimetemp++
If addtimetemp = 2
{
Addtime = 2
Addtimetemp = 0
Return
}
}
If addtime = 2
{
Addtimetemp++
If addtimetemp = 2
{
Refreshrate= 10
Addtimetemp = 0
addtime = 3
SplashTextOn,350,50,Refresh time reduced, Macro will refresh ACM every 10 Serials to help reduce the amount of timeouts from a slow ACM
Sleep 5000
SplashTextOff
Return
}
}
If addtime = 3
{
Return
}
Return
}
$^1::
{
Gosub, FormatSerialsMacro
return
}
FormatSerialsMacro:
{
Clipboard5 =
FullString =
oddballvar =
Checkers =
Clipboard =
Clipboard1 =
Checkeend  =
Editfield =
Parseclip =
Guicontrol,, Editfield, %Editfield%
newline = `n
formatfound =
Stringthree =
Editfield =
FullString =
oddballvar =
Checkers =
clipboard =
sleep 100
Send ^c
sleep 200
if clipboard =
{
Msgbox, Error. Please ensure that text is selected before pressing Ctrl + 1.
Exit
}
FullString = %clipboard%`,
oddballvar = %Fullstring%
Checkers = %Fullstring%
StringReplace, Checkers,Checkers,`n,,All
StringReplace, Checkers,Checkers,`r,,All
StringReplace, Checkers,Checkers,%A_Space%,,All
StringReplace, Checkers,Checkers, `),`)`n,All
StringReplace, Checkers,Checkers, 1`-up,,All
StringReplace, Checkers,Checkers,  `-up,`-99999, All
StringReplace, Checkers,Checkers, and,,All
Loop, Parse, Checkers, `r`n
{
StringGetPos, pos, A_loopfield, :, 1
String%Counter% := SubStr(A_LoopField, pos+2)
StringTemp := String%Counter%
StringReplace, newStr,StringTemp,%A_Space%,,All
StringReplace, NewStr2,NewStr, `),`,,All
StringReplace, NewStr3,NewStr2, 1`-up,,All
StringReplace, NewStr4,NewStr3, `-up,`-99999, All
StringReplace, NewStr5,NewStr4, `,,`,`n, All
Counter++
If Newstr5 =
{
Clipboard1 =  %clipboard1%%NewStr5%
}
Else
{
Clipboard1 =  %clipboard1%%NewStr5%
}
}
Clipboard1 = %clipboard1%`,
StringReplace, Clipboard1,Clipboard1, `,,,All
Counter = 0
Parseclip = %clipboard1%
Loop, parse, Parseclip, `n
{
If a_loopfield =
{
Continue
}
TotalPrefixes++
Counter++
addcomma = %A_LoopField%`,
StringGetPos, pos, addcomma, `,, 1
If pos = 3
{
StringReplace, Clippyt,addcomma,`,,00001`-99999`,%newline%,all
Clipboard5 = %Clipboard5%%Clippyt%
}
Else if pos != 3
{
StringMid, String%Counter%prefix, A_loopfield, 1, 3
Prefixstore := String%Counter%prefix
StringTrimLeft, String%Counter%left, A_loopfield, 3
StringTemppre := String%Counter%left
Gosub, add_digits
Stringtemp = %Prefixstore%%Stringtempend%
StringReplace, StringTemp,StringTemp,`),`,`n,all
Clippyt = %Stringtemp%
Clipboard5 = %Clipboard5%%Clippyt%
}
}
Editfieldbreaks = %Clipboard5%%newline%
CLipboard6 = %Clipboard5%
Gosub, Combineserials
Prefixcombinecount = 0
Loop, Parse, Prefixmatching, `, all
{
if a_loopfield =
Continue
Prefixcombinecount++
finalcombine := Prefix%A_LoopField%
CLipboard7 = %clipboard7%%finalcombine%`,%newline%
}
EditfieldCombine = %clipboard7%
gosub, SerialbreakquestionGUI
If combineser = 0
{
Editfield = %EditFieldbreaks%
Guicontrol,1:, Editfield, %EditFieldbreaks%
}
If combineser = 1
{
Editfield = %EditfieldCombine%%newline%
totalprefixes = %Prefixcombinecount%
Guicontrol,1:, Editfield, %EditfieldCombine%%newline%
}
Guicontrol,, reloadprefixtext, There are a total of %totalprefixes% Serial Numbers to add to ACM
Winactivate, Enter Serial Macro
Guicontrol, Focus, Editfield
Send {Ctrl Down}{End}{Ctrl Up}
Sleep 300
Send {BackSpace}{*}{*}{*}
send {Ctrl Down}{Home}{Ctrl Up}
gosub, Radio1h
Gui, Submit, NoHide
Return
}
add_digits:
{
Stringtempend =
combinestring =
Loop, Parse, StringTemppre, `-
{
StringReplace, StringTemprep,A_LoopField,`),,
int = %StringTemprep%
Loop, % 5-StrLen(int)
int = 0%int%
combinestring = %combinestring%`-%Int%
}
combinestring := SubStr(combinestring, 2)
Stringtempend = %combinestring%`)
return
}
combineserials:
{
counting = 0
Match = 0
Loop, Parse, clipboard6, `n,all
{
counting++
StringMid, String%Counting%prefix, A_loopfield, 1, 3
Checkprefix := String%Counting%prefix
If checkprefix = `,
{
checkprefix =
lastnums =
Continue
}
gosub, matchprefix
Stringmid, String%Counting%number,A_LoopField,4,5
Stringmid, String%Counting%check,A_LoopField,9,1
endchar :=  String%Counting%check
begnumcheck := String%Counting%number
If endchar = `-
{
Stringmid, String%Counting%last,A_LoopField,10,5
Lastnums :=  String%Counting%last
}
If Endchar = `,
{
Endchar = `-
lastnums = %begnumcheck%
}
If Match = 1
{
Gosub, Checkvalues
Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums%
Serialnumzz := Prefix%Checkprefix%
Match = 0
Continue
}
Else if Match = 0
{
Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums%
}
}
Return
}
Checkvalues:
{
oldprefix := Prefix%Checkprefix%
Stringmid, prefixbeg,oldprefix,4,5
Stringmid, prefixmid,oldprefix,9,1
endcharcheck :=  prefixmid
Stringmid, Prefixlast,oldprefix,10,5
Lastnumcheck :=  Prefixlast
If 	prefixbeg > %lastnums%
{
lastnums = %prefixbeg%
}
If 	prefixbeg < %begnumcheck%
{
begnumcheck = %prefixbeg%
}
If 	lastnumcheck > %lastnums%
{
lastnums = %lastnumcheck%
}
return
}
matchprefix:
{
Loop, Parse, Prefixmatching,`,,,all
{
If a_loopfield =
{
blank = 1
Continue
}
IF A_LoopField = %CHeckprefix%
{
Match = 1
Return
}
}
If match = 1
{
return
}
If blank = 1
{
Blank = 0
}
If match != 1
{
Blank = 0
match = 0
Prefixmatching = %Prefixmatching%%CHeckprefix%`,
}
return
}
aboutmacro:
{
Gui, 7:Add, Picture, x0 y0 w525 h300 +0x4000000, C:\SerialMacro\background.png
gui, 7:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, This program was designed to help increase the speed and accuracy of entering machine serial prefixes form the Pubtool into ACM.
gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans, To get the latest version, go to the File menu and select Check for updates.
gui, 7:add, Text, xp yp+25 w300 h20 BackgroundTrans,  The location of the macro is at the following box account:
gui 7:font, CBlue Underline
gui, 7:add, Text, xp+275 yp w500 h20 BackgroundTrans gboxlink, https://cat.box.com/v/EnterSerialMacro
gui 7:font,
gui, 7:add, Text, xp-275 yp+25 w500 h40 BackgroundTrans , For reporting bugs or enhancement requests, Please send an email with the Subject line "Enter Serial Macro" to the below address
gui 7:font, CBlue Underline
gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans gemaillink, Karnia_Jarett_S@cat.com
gui 7:font,
gui, 7:add, Text, xp+300 yp+45 w500 h20 BackgroundTrans, This program was created by
gui, 7:add, Text, xp yp+25 w500 h20 BackgroundTrans, and is maintained by Jarett Karnia
Gui, 7:Show, x100 y100 w525 h300, About -- Enter Serial Macro V%Version_Number%
return
}
emaillink:
{
Run,  mailto:Karnia_Jarett_S@cat?Subject=Enter Serials Macro
return
}
boxlink:
{
Run, https://cat.box.com/v/EnterSerialMacro
return
}
SerialbreakquestionGUI:
{
gui 1: -alwaysontop
Gui, 8:Add, Picture, x0 y0 w400 h75 +0x4000000, C:\SerialMacro\background.png
Gui, 8: Add, text, x10 y20 BackgroundTrans, Do you want to combine the serial breaks, or keep the serial breaks seperated?
Gui, 8:add, button, xp+50 yp+20 gcombinequstion, Combine
gui, 8:add, button, xp+170 yp gkeepseperated, Keep Seperated
Gui, 8:show, w400 h75
gui 8: +alwaysontop
Pause, on
return
}
combinequstion:
{
Pause, off
gui 1: +alwaysontop
Gui, 8:destroy
combineser = 1
GOsub, radio1h
Return
}
keepseperated:
{
Pause, off
gui 1: +alwaysontop
Gui, 8:destroy
combineser = 0
GOsub, radio1h
Return
}
SerialsGUIscreen:
{
Menu, BBBB, Add, &Check For Update , Versioncheck
Menu, BBBB, Add, &Options, OptionsGui
Menu, BBBB, Add,
Menu, CCCC, Add, &Run							(Crtl + 2), Enterallserials
Menu, CCCC, Add, &Pause/Unpause 				(Pause / Insert), pausesub
Menu, CCCC, Add, &Stop Macro					(ESC), Exitprogram
Menu, CCCC, Add, &Reload Macro, restartmacro
Menu, BBBB, Add, &Exit							(Ctrl + Q), Quitapp
Menu, DDDD, Add, &How To Use					(F1), HowTo
Menu, DDDD, Add, &About , Aboutmacro
Menu, MyMenuBar, Add, &File, :BBBB
Menu, MyMenuBar, Add, &Macro, :CCCC
Menu, MyMenuBar, Add, &Help, :DDDD
SplashTextOff
If totalprefixes < 1
{
TotalPrefixzero = 0
}
Else
TotalPrefixzero = %totalprefixes%
height := A_ScreenHeight/2 - 300
width := A_ScreenWidth/2
TT:=TT("GUI=1 Initial=1500 AUTOPOP=3000")
gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfield%
tt.add("Edit1" ,"This area holds the formatted serials after you select the serials in the pubtool and press Ctrl + 1")
gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2%
Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstartmacro , C:\SerialMacro\Start.png
Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\Running.png
Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\Paused.png
Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestartmacro, C:\SerialMacro\Stopped.png
Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\background.png
tt.add("Edit2","This area holds the serials that were entered into ACM. This field will also show is a serial was not added and the reason that it wasn’t.")
Gui 1:Add, Edit, xp+145 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%
tt.add("edit3","This area holds the next serial what will be added to ACM")
Gui 1:Add, Text, x5 y5 w300 h25 BackgroundTrans +Center vreloadprefixtext, There are a total of %totalprefixzero% Serial Numbers to add to ACM
tt.add("Text1","This is the amount of Serial Prefixes that will be added to ACM with this macro")
Gui 1:add, Radio, xp+25 yp+25 w120 h20 BackGroundTrans vradio1 gradio1h, Serials to be added:
tt.add("radio1","Click on this button to see the Serials that will be added into ACM")
Gui 1:add, Radio, xp+155 yp w130 h20 BackGroundTrans vradio2 gradio2h, Serials already added
tt.add("radio2","Click on this button to see the Serials that were added into ACM")
Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans vserialsentered, Number of Serials successfully added to ACM = %Serialcount%
tt.add("Text2","This is the amount of Serial Prefixes was successfully added to ACM with this macro")
Gui 1:Add, Text, Xp Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
Gui 1:Add, Text, xp yp+15 w250 h13  BackGroundTrans , Or press Pause Button on keyboard to Pause macro, Press Pause again to resume macro.
Gui 1:Add, Text, xp yp+20 w135 h13  BackgroundTrans , Next Serial to add To ACM:
Gui 1:Add, button, xp yp+25 BackGroundTrans grerunclicks, Rerun Shift + Click positions
Gui 1:Menu, MyMenuBar
Gui 1:Show,  x%width% y%height% , Enter Serial Macro V%Version_Number%
gui 1: +alwaysontop
Guicontrol,hide, Editfield2,
Guicontrol,show, Editfield,
Guicontrol, Focus, Editfield
Gui 1:Submit, NoHide
Return
}
OptionsGui:
{
gui 10: +alwaysontop
Gui , 1: -AlwaysOnTop
gui 10:add, text, xp yp+25 w250 h20 ,Refreash ACM Rate (After how many entered serials)
gui 10:add, edit, xp+251 yp-1 w30 veditfield5 , %refreshrate%
gui 10:add, button, xp-251 yp+26 h20 w75 gsavesets, Save Settings
Gui, 10:Add, Picture, x0 y0 w325 h100 +0x4000000 , C:\SerialMacro\background.png
gui 10:show, w325 h100
Guicontrol,10:, editfield5, %refreshrate%
gui, 10:submit, nohide
return
}
savesets:
{
Gui , 1: +AlwaysOnTop
GuiControlGet,Refreshrate,,editfield5
IniWrite, %refreshrate%,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
IniWrite, %mouseposclicks%,  C:\Serialmacro\Settings.ini,searches,searchclick
IniWrite, %imagesearchoption%,  C:\Serialmacro\Settings.ini,searches,searchimg
gui 10:submit, nohide
gui 10:destroy
return
}
10guiclose:
{
Gui , 1: +AlwaysOnTop
gui 10:destroy
return
}
mouseclickss:
{
Entermethod = 2
return
}
imagessearching:
{
Entermethod = 1
return
}
startmacro:
{
skipbox = 0
gosub, enterallserials
return
}
rerunclicks:
{
Stop = 1
Send {shift down}{Shift up}
Return
}
pauses:
{
Pause, off
return
}
radio1h:
{
GuiControl,, radio2, 0
GuiControl,, radio1,1
Guicontrol,show, Editfield,
Guicontrol,hide, Editfield2,
gui, submit, nohide
return
}
radio2h:
{
GuiControl,, radio1,0
GuiControl,, radio2,1
Guicontrol,show, Editfield2,
Guicontrol,hide, Editfield,
gui, submit, nohide
return
}
radio3h:
{
GuiControl,, radio1,0
GuiControl,, radio2,0
Guicontrol,hide, Editfield,
Guicontrol,hide, Editfield2,
return
}
Guiclose:
{
msgbox,262148,Quit -- Enter Serial Macro V%Version_Number%, Are you sure you want to quit?
ifMsgBox Yes
{
ExitApp
}
Else
{
Return
}
Return
}
restartmacro:
{
msgbox,262148,Reload -- Enter Serial Macro V%Version_Number%, Are you sure that you want to reload the program?
IfMsgBox yes
{
Reload
}
Else
{
return
}
return
}
Macrotimedout:
{
Gui, 3:Add, Picture, x0 y0 w525 h300 +0x4000000, C:\SerialMacro\background.png
gui, -alwaysontop
gui, 3:add, Text, xp+5 yp+5 w500 h20 BackgroundTrans, Macro is now paused due to not being able to find an empty prefix box in 5 Seconds.
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans, Press Pause to unpause after you press OK on this box and fixed the issue that is causing ACM to lockup.
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans,  The delay is most likely due to one of the following:
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans,  - The Eng Model needs to be manually added with the drop down menu.
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans,  - Serial # Does not Exist in ACM
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans,  - ACM is running slow
gui, 3:add, Text, xp yp+25 w500 h20 BackgroundTrans,  - Something really bad happened and the effectivity page is not there anymore.
gui, 3:add, Text, xp yp+25 w500 h40 BackgroundTrans,  Please Press the one of the buttons below after you Idenitify the cause. This action will reflect in the Serials Added Screen.
gui, 3:add, Button, xp+1 yp+50 w100 h50 BackgroundTrans gEngmodel,  More than one Engineering Model
gui, 3:add, Button, xp+125 yp w100 h50 BackgroundTrans gSerialnogo,  Serial Does Not Exist
gui, 3:add, Button, xp+125 yp w100 h50 BackgroundTrans gacmlong,  ACM Took to long
gui, 3:add, Button, xp+125 yp w100 h50 BackgroundTrans  gAnarchy,  Something else happened :(
Gui, 3:Show, w525 h300, Macro Timed Out -- Enter Serial Macro V%Version_Number%
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Pause, on
return
}
#IfWinActive Enter Serial Macro
F1::
{
gosub, howto
Return
}
#IfWinActive
Howto:
{
splashtexton,,Enter Serial Macro, Loading PDF
Run, C:\SerialMacro\How to use Serial Macros.pdf
sleep 2000
SplashTextOff
return
}
TT_Init(){
global _TOOLINFO:="cbSize,uFlags,PTR hwnd,PTR uId,_RECT rect,PTR hinst,LPTSTR lpszText,PTR lParam,void *lpReserved"
,_RECT:="left,top,right,bottom"
,_NMHDR:="HWND hwndFrom,UINT_PTR idFrom,UINT code"
,_NMTVGETINFOTIP:="_NMHDR hdr,UINT uFlags,UInt link"
,_CURSORINFO:="cbSize,flags,HCURSOR hCursor,x,y"
,_ICONINFO:="fIcon,xHotspot,yHotSpot,HBITMAP hbmMask,HBITMAP hbmColor"
,_BITMAP:="LONG bmType,LONG bmWidth,LONG bmHeight,LONG bmWidthBytes,WORD bmPlanes,WORD bmBitsPixel,LPVOID bmBits"
,_SHFILEINFO:="HICON hIcon,iIcon,DWORD dwAttributes,TCHAR szDisplayName[260],TCHAR szTypeName[80]"
,_TBBUTTON:="iBitmap,idCommand,BYTE fsState,BYTE fsStyle,BYTE bReserved[" (A_PtrSize=8?6:2) "],DWORD_PTR dwData,INT_PTR iString"
static _:={base:{__Delete:"TT_Delete"}}
return _
}
TT(options:="",text:="",title:=""){
global _RECT,_TOOLINFO
static HWND_TOPMOST:=-1,SWP_NOMOVE:=0x2, SWP_NOSIZE:=0x1, SWP_NOACTIVATE:=0x10
static _:=TT_Init(),base:={Color:"TT_Color",Show:"TT_Show",Hide:"TTM_Trackactivate",Close:"TT_Close",Add:"TT_Add",AddTool:"TTM_AddTool"
,Del:"TT_Del",Title:"TTM_SetTitle",Text:"TT_Text",ACTIVATE:"TTM_ACTIVATE",Set:"TT_Set"
,ADDTOOL:"TTM_ADDTOOL",Remove:"TT_Remove",Icon:"TT_Icon",Font:"TT_Font",ADJUSTRECT:"TTM_ADJUSTRECT"
,DELTOOL:"TTM_DELTOOL",ENUMTOOLS:"TTM_ENUMTOOLS",GETBUBBLESIZE:"TTM_GETBUBBLESIZE"
,GETCURRENTTOOL:"TTM_GETCURRENTTOOL",GETDELAYTIME:"TTM_GETDELAYTIME",GETMARGIN:"TTM_GETMARGIN"
,GETMAXTIPWIDTH:"TTM_GETMAXTIPWIDTH",GETTEXT:"TTM_GETTEXT",GETTIPBKCOLOR:"TTM_GETTIPBKCOLOR"
,GETTIPTEXTCOLOR:"TTM_GETTIPTEXTCOLOR",GETTITLE:"TTM_GETTITLE",GETTOOLCOUNT:"TTM_GETTOOLCOUNT"
,GETTOOLINFO:"TTM_GETTOOLINFO",HITTEST:"TTM_HITTEST",NEWTOOLRECT:"TTM_NEWTOOLRECT",POP:"TTM_POP"
,POPUP:"TTM_POPUP",RELAYEVENT:"TTM_RELAYEVENT",SETDELAYTIME:"TTM_SETDELAYTIME",SETMARGIN:"TTM_SETMARGIN"
,SETMAXTIPWIDTH:"TTM_SETMAXTIPWIDTH",SETTIPBKCOLOR:"TTM_SETTIPBKCOLOR",SETTIPTEXTCOLOR:"TTM_SETTIPTEXTCOLOR"
,SETTITLE:"TTM_SETTITLE",SETTOOLINFO:"TTM_SETTOOLINFO",SETWINDOWTHEME:"TTM_SETWINDOWTHEME"
,TRACKACTIVATE:"TTM_TRACKACTIVATE",TRACKPOSITION:"TTM_TRACKPOSITION",UPDATE:"TTM_UPDATE"
,UPDATETIPTEXT:"TTM_UPDATETIPTEXT",WINDOWFROMPOINT:"TTM_WINDOWFROMPOINT"
,"base":{__Call:"TT_Set"}}
Parent:="",Gui:="",ClickTrough:="",Style:="",NOFADE:="",NoAnimate:="",NOPREFIX:="",AlwaysTip:="",ParseLinks:="",CloseButton:="",Balloon:="",maxwidth:=""
,INITIAL:="",AUTOPOP:="",RESHOW:="",OnClick:="",OnClose:="",OnShow:="",ClickHide:="",HWND:="",Center:="",RTL:="",SUB:="",Track:="",Absolute:=""
,TRANSPARENT:="",Color:="",Background:="",icon:=0
if (options+0!="")
Parent:=options
else If (options){
Loop,Parse,options,%A_Space%,%A_Space%
If (istext){
If (SubStr(A_LoopField,0)="'")
%istext%:=string A_Space SubStr(A_LoopField,1,StrLen(A_LoopField)-1),istext:="",string:=""
else
string.= A_Space A_LoopField
} else If (A_LoopField ~= "i)AUTOPOP|INITIAL|PARENT|RESHOW|MAXWIDTH|ICON|Color|BackGround|OnClose|OnClick|OnShow|GUI|NOPREFIX|TRACK")
{
RegExMatch(A_LoopField,"^(\w+)=?(.*)?$",option)
If ((Asc(option2)=39 && SubStr(A_LoopField,0)!="'") && (istext:=option1) && (string:=SubStr(option2,2)))
Continue
%option1%:=option2
} else if ( option2:=InStr(A_LoopField,"=")){
option1:=SubStr(A_LoopField,1,option2-1)
%option1%:=SubStr(A_LoopField,option2+1)
} else if (A_LoopField)
%A_LoopField% := 1
}
If (Parent && Parent<100 && !DllCall("IsWindow","PTR",Parent)){
Gui,%Parent%:+LastFound
Parent:=WinExist()
} else if (GUI){
Gui, %GUI%:+LastFound
Parent:=WinExist()
} else if (Parent=""){
Parent:=A_ScriptHwnd
}
T:=Object("base",base)
,T.HWND := DllCall("CreateWindowEx", "UInt", (ClickTrough?0x20:0)|0x8, "str", "tooltips_class32", "PTR", 0
, "UInt",0x80000000|(Style?0x100:0)|(NOFADE?0x20:0)|(NoAnimate?0x10:0)|((NOPREFIX+1)?(NOPREFIX?0x2:0x2):0x2)|(AlwaysTip?0x1:0)|(ParseLinks?0x1000:0)|(CloseButton?0x80:0)|(Balloon?0x40:0)
, "int",0x80000000,"int",0x80000000,"int",0x80000000,"int",0x80000000, "PTR",Parent?Parent:0,"PTR",0,"PTR",0,"PTR",0,"PTR")
,DllCall("SetWindowPos","PTR",T.HWND,"PTR",HWND_TOPMOST,"Int",0,"Int",0,"Int",0,"Int",0
,"UInt",SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE)
,_.Insert(T)
,T.SETMAXTIPWIDTH(MAXWIDTH?MAXWIDTH:A_ScreenWidth)
If !(AUTOPOP INITIAL RESHOW)
T.SETDELAYTIME()
else T.SETDELAYTIME(2,AUTOPOP?AUTOPOP:-1),T.SETDELAYTIME(3,INITIAL?INITIAL:-1),T.SETDELAYTIME(1,RESHOW?RESHOW:-1)
T.fulltext:=text,T.maintext:=RegExReplace(text,"<a\K[^<]*?>",">")
If (OnClick)
ParseLinks:=1
T.rc:=Struct(_RECT)
,T.P:=Struct(_TOOLINFO),P:=T.P,P.cbSize:=sizeof(_TOOLINFO)
,P.uFlags:=(HWND?0x1:0)|(Center?0x2:0)|(RTL?0x4:0)|(SUB?0x10:0)|(Track+1?(Track?0x20:0):0x20)|(Absolute?0x80:0)|(TRANSPARENT?0x100:0)|(ParseLinks?0x1000:0)
,P.hwnd:=Parent
,P.uId:=Parent
,P.lpszText[""]:=T.GetAddress("maintext")?T.GetAddress("maintext"):0
,T.AddTool(P[])
If (Theme)
T.SETWINDOWTHEME()
If (Color)
T.SETTIPTEXTCOLOR(Color)
If (Background)
T.SETTIPBKCOLOR(BackGround)
T.SetTitle(T.maintitle:=title,icon)
If ((T.OnClick:=OnClick)||(T.OnClose:=OnClose)||(T.OnShow:=OnShow))
T.OnClose:=OnClose,T.OnShow:=OnShow,T.ClickHide:=ClickHide,OnMessage(0x4e,"TT_OnMessage")
Return T
}
TT_Delete(this){
Loop % this.MaxIndex()
{
this[i:=A_Index].DelTool(this[i].P[])
,DllCall("DestroyWindow","PTR",this[i].HWND)
for id,tool in this[i].T
this[i].DelTool(tool[])
this.Remove(i)
}
TT_GetIcon()
}
TT_Remove(T:=""){
static _:=TT_Init()
for id,Tool in _
{
If (T=Tool){
_[id]:=_[_.MaxIndex()],_.Remove(id)
for id,tools in Tool.T
Tool.DelTool(tools[])
Tool.DelTool(Tool.P[])
,DllCall("DestroyWindow","PTR",Tool.HWND)
break
}
}
}
TT_OnMessage(wParam,lParam,msg,hwnd){
global _NMTVGETINFOTIP,_NMHDR
static TTN_FIRST:=0xfffffdf8, _:=TT_Init()
,HDR:=Struct(_NMTVGETINFOTIP)
HDR[]:=lParam
If !InStr(".1.2.3.","." (m:= TTN_FIRST-HDR.hdr.code) ".")
Return
p:=HDR.hdr.hwndFrom
for id,T in _
If (p=T.hwnd)
break
for id,object in _
If (p=object.hwnd && T:=object)
break
text:=T.fulltext
If (m=1){
If IsFunc(T.OnShow)
T.OnShow(T,"")
} else If (m=2){
If IsFunc(T.OnClose)
T.OnClose(T,"")
T.TRACKACTIVATE(0,T.P[])
} else If InStr(text,"<a"){
If (T.ClickHide)
T.TRACKACTIVATE(0,T.P[])
If (SubStr(LTrim(text:=SubStr(text,InStr(text,"<a",0,1,HDR.link+1)+2)),1,1)=">")
action:=SubStr(text,InStr(text,">")+1,InStr(text,"</a>")-InStr(text,">")-1)
else action:=Trim(action:=SubStr(text,1,InStr(text,">")-1))
If IsFunc(func:=T.OnClick)
T.OnClick(T,action)
}
Return true
}
TT_ADD(T,Control,Text:="",uFlags:="",Parent:=""){
Global _TOOLINFO
DetectHiddenWindows:=A_DetectHiddenWindows
DetectHiddenWindows,On
if (Parent){
If (Parent && Parent<100 and !DllCall("IsWindow","PTR",Parent)){
Gui %Parent%:+LastFound
Parent:=WinExist()
}
T["T",Abs(Parent)]:=Tool:=Struct(_TOOLINFO)
,Tool.uId:=Parent,Tool.hwnd:=Parent,Tool.uFlags:=(0|16)
,DllCall("GetClientRect","PTR",T.HWND,"PTR", T[Abs(Parent)].rect[])
,T.ADDTOOL(T["T",Abs(Parent)][])
}
If (text="")
ControlGetText,text,%Control%,% "ahk_id " (Parent?Parent:T.P.hwnd)
If (Control+0="")
ControlGet,Control,Hwnd,,%Control%,% "ahk_id " (Parent?Parent:T.P.hwnd)
If (uFlags)
If (uFlags+0="")
{
Loop,Parse,uflags,%A_Space%,%A_Space%
If (A_LoopField)
%A_LoopField% := 1
uFlags:=(HWND?0x1:HWND=""?0x1:0)|(Center?0x2:0)|(RTL?0x4:0)|(SUB?0x10:0)|(Track?0x20:0)|(Absolute?0x80:0)|(TRANSPARENT?0x100:0)|(ParseLinks?0x1000:0)
}
Tool:=T["T",Abs(Control)]:=Struct(_TOOLINFO)
,Tool.cbSize:=sizeof(_TOOLINFO)
,T[Abs(Control),"text"]:=RegExReplace(text,"<a\K[^<]*?>",">")
,Tool.uId:=Control,Tool.hwnd:=Parent?Parent:T.P.hwnd,Tool.uFlags:=uFlags?(uFlags|16):(1|16)
,Tool.lpszText[""]:=T[Abs(Control)].GetAddress("text")
,DllCall("GetClientRect","PTR",T.HWND,"PTR",Tool.rect[])
T.ADDTOOL(Tool[])
DetectHiddenWindows,%DetectHiddenWindows%
}
TT_DEL(T,Control){
If (!Control)
Return 0
If (Control+0="")
ControlGet,Control,Hwnd,,%Control%,% "ahk_id " t.P.hwnd
T.DELTOOL(T.T[Abs(Control)][]),T.T.Remove(Abs(Control))
}
TT_Color(T,Color:="",Background:=""){
static TTM_SETTIPBKCOLOR:=0x413,TTM_SETTIPTEXTCOLOR:=0x414
,Black:=0x000000,Green:=0x008000,Silver:=0xC0C0C0,Lime:=0x00FF00,Gray:=0x808080,Olive:=0x808000
,White:=0xFFFFFF,Yellow:=0xFFFF00,Maroon:=0x800000,Navy:=0x000080,Red:=0xFF0000,Blue:=0x0000FF
,Purple:=0x800080,Teal:=0x008080,Fuchsia:=0xFF00FF,Aqua:=0x00FFFF
If (Color!="")
T.SETTIPTEXTCOLOR(Color)
If (BackGround!="")
T.SETTIPBKCOLOR(BackGround)
}
TT_Text(T,text){
static TTM_UPDATETIPTEXT:=0x400+(A_IsUnicode?57:12),TTM_UPDATE:=0x400+29
T.fulltext:=text,T.maintext:=RegExReplace(text,"<a\K[^<]*?>",">"),T.P.lpszText[""]:=text!=""?T.GetAddress("maintext"):0
,T.UPDATETIPTEXT()
}
TT_Icon(T,icon:=0,icon_:=1,default:=1){
static TTM_SETTITLE := 0x400 + (A_IsUnicode ? 33 : 32)
If icon
If (icon+0="")
If !icon:=TT_GetIcon(icon,icon_)
icon:=default
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETTITLE,"PTR",icon+0,"PTR",T.GetAddress("maintitle"),"PTR"),T.UPDATE()
}
TT_GetIcon(File:="",Icon_:=1){
global _ICONINFO,_SHFILEINFO
static hIcon:={},AW:=A_IsUnicode?"W":"A",pToken:=0
,temp1:=DllCall( "LoadLibrary", "Str","gdiplus","PTR"),temp2:=VarSetCapacity(si, 16, 0) (si := Chr(1)) DllCall("gdiplus\GdiplusStartup", "PTR*",pToken, "PTR",&si, "PTR",0)
static sfi:=Struct(_SHFILEINFO),sfi_size:=sizeof(_SHFILEINFO),SmallIconSize:=DllCall("GetSystemMetrics","Int",49)
If !File {
for file,obj in hIcon
If IsObject(obj){
for icon,handle in obj
DllCall("DestroyIcon","PTR",handle)
} else
DllCall("DestroyIcon","PTR",handle)
Return
}
If (CR:=InStr(File,"`r") || LF:=InStr(File,"`n"))
File:=SubStr(file,1,CR<LF?CR-1:LF-1)
If (hIcon[File,Icon_])
Return hIcon[file,Icon_]
else if (hIcon[File] && !IsObject(hIcon[File]))
return hIcon[File]
SplitPath,File,,,Ext
if (hIcon[Ext] && !IsObject(hIcon[Ext]))
return hIcon[Ext]
else If (ext = "cur")
Return hIcon[file,Icon_]:=DllCall("LoadImageW", "PTR", 0, "str", File, "uint", ext="cur"?2:1, "int", 0, "int", 0, "uint", 0x10,"PTR")
else if InStr(",EXE,ICO,DLL,LNK,","," Ext ","){
If (ext="LNK"){
FileGetShortcut,%File%,Fileto,,,,FileIcon,FileIcon_
File:=!FileIcon ? FileTo : FileIcon
}
SplitPath,File,,,Ext
DllCall("PrivateExtractIcons", "Str", File, "Int", Icon_-1, "Int", SmallIconSize, "Int", SmallIconSize, "PTR*", Icon, "PTR*", 0, "UInt", 1, "UInt", 0, "Int")
Return hIcon[File,Icon_]:=Icon
} else if (Icon_=""){
If !FileExist(File){
if RegExMatch(File,"^[0-9A-Fa-f]+$")
{
nSize := StrLen(File)//2
VarSetCapacity( Buffer,nSize )
Loop % nSize
NumPut( "0x" . SubStr(File,2*A_Index-1,2), Buffer, A_Index-1, "Char" )
} else Return
} else {
FileGetSize,nSize,%file%
FileRead,Buffer,*c %file%
}
hData := DllCall("GlobalAlloc", "UInt",2, "UInt", nSize,"PTR")
,pData := DllCall("GlobalLock", "PTR",hData,"PTR")
,DllCall( "RtlMoveMemory", "PTR",pData, "PTR",&Buffer, "UInt",nSize )
,DllCall( "GlobalUnlock", "PTR",hData )
,DllCall( "ole32\CreateStreamOnHGlobal", "PTR",hData, "Int",True, "PTR*",pStream )
,DllCall( "gdiplus\GdipCreateBitmapFromStream", "UInt",pStream, "PTR*",pBitmap )
,DllCall( "gdiplus\GdipCreateHBITMAPFromBitmap", "PTR",pBitmap, "PTR*",hBitmap, "UInt",0 )
,DllCall( "gdiplus\GdipDisposeImage", "PTR",pBitmap )
,ii:=Struct(_ICONINFO)
,ii.ficon:=1,ii.hbmMask:=hBitmap,ii.hbmColor:=hBitmap
return hIcon[File]:=DllCall("CreateIconIndirect","PTR",ii[],"PTR")
} else If DllCall("Shell32\SHGetFileInfo" AW, "str", File, "uint", 0, "PTR", sfi[], "uint", sfi_size, "uint", 0x101,"PTR")
Return hIcon[Ext] := sfi.hIcon
}
TT_Close(T){
T.text("")
}
TT_Show(T,text:="",x:="",y:="",title:="",icon:=0,icon_:=1,defaulticon:=1){
global _TBBUTTON,_BITMAP,_ICONINFO,_CURSORINFO,_RECT
static pcursor:= Struct(_CURSORINFO),init:=(pcursor.cbSize:=sizeof(_CURSORINFO))
,picon:=Struct(_ICONINFO) ,pbitmap:=Struct(_BITMAP)
,TB:=Struct(_TBBUTTON) ,RC:=Struct(_RECT)
xo:=0,xs:=0,yo:=0,ys:=0
If (Text!="")
T.Text(text)
If (title!="")
T.SETTITLE(title,icon,icon_,defaulticon)
If (x="TrayIcon" || y="TrayIcon"){
DetectHiddenWindows,% (DetectHiddenWindows:=A_DetectHiddenWindows ? "On" : "On")
PID:=DllCall("GetCurrentProcessId")
hWndTray:=WinExist("ahk_class Shell_TrayWnd")
ControlGet,hWndToolBar,Hwnd,,ToolbarWindow321,ahk_id %hWndTray%
WinGet, procpid, Pid, ahk_id %hWndToolBar%
DataH   := DllCall( "OpenProcess", "uint", 0x38, "int", 0, "uint", procpid,"PTR")
,bufAdr  := DllCall( "VirtualAllocEx", "PTR", DataH, "PTR", 0, "uint", sizeof(_TBBUTTON), "uint", MEM_COMMIT:=0x1000, "uint", PAGE_READWRITE:=0x4,"PTR")
Loop % max:=DllCall("SendMessage","PTR",hWndToolBar,"UInt",0x418,"PTR",0,"PTR",0,"PTR")
{
i:=max-A_Index
DllCall("SendMessage","PTR",hWndToolBar,"UInt",0x417,"PTR",i,"PTR",bufAdr,"PTR")
,DllCall("ReadProcessMemory", "PTR", DataH, "PTR", bufAdr, "PTR", TB[], "ptr", sizeof(TB), "ptr", 0)
,DllCall("ReadProcessMemory", "PTR", DataH, "PTR", TB.dwData, "PTR", RC[], "PTR", 8, "PTR", 0)
WinGet,BWPID,PID,% "ahk_id " NumGet(RC[],0,"PTR")
If (BWPID!=PID)
continue
If (TB.fsState>7){
ControlGetPos,xc,yc,xw,yw,Button2,ahk_id %hWndTray%
xc+=xw/2, yc+=yw/4
} else {
ControlGetPos,xc,yc,,,ToolbarWindow321,ahk_id %hWndTray%
DllCall("SendMessage","PTR",hWndToolBar,"UInt",0x41d,"PTR",i,"PTR",bufAdr,"PTR")
,DllCall( "ReadProcessMemory", "PTR", DataH, "PTR", bufAdr, "PTR", RC[], "PTR", sizeof(RC), "PTR", 0 )
,halfsize:=RC.bottom/2
,xc+=RC.left + halfsize
,yc+=RC.top + (halfsize/1.5)
}
WinGetPos,xw,yw,,,ahk_id %hWndTray%
xc+=xw,yc+=yw
break
}
If (!xc && !yc){
If (A_OsVersion~="i)Win_7|WIN_VISTA")
ControlGetPos,xc,yc,xw,yw,Button1,ahk_id %hWndTray%
else
ControlGetPos,xc,yc,xw,yw,Button2,ahk_id %hWndTray%
xc+=xw/2, yc+=yw/4
WinGetPos,xw,yw,,,ahk_id %hWndTray%
xc+=xw,yc+=yw
}
DllCall( "VirtualFreeEx", "PTR", DataH, "PTR", bufAdr, "PTR", 0, "uint", MEM_RELEASE:=0x8000)
,DllCall( "CloseHandle", "PTR", DataH )
DetectHiddenWindows % DetectHiddenWindows
If (x="TrayIcon")
x:=xc
If (y="TrayIcon")
y:=yc
}
If (x ="" || y =""){
pCursor.cbSize:=sizeof(pCursor)
,DllCall("GetCursorInfo", "ptr", pCursor[])
,DllCall("GetIconInfo", "ptr", pCursor.hCursor, "ptr", pIcon[])
If (picon.hbmColor)
DllCall("DeleteObject", "ptr", pIcon.hbmColor)
DllCall("GetObject", "ptr", pIcon.hbmMask, "uint", sizeof(_BITMAP), "ptr", pBitmap[])
,hbmo := DllCall("SelectObject", "ptr", cdc:=DllCall("CreateCompatibleDC", "ptr", sdc:=DllCall("GetDC","ptr",0,"ptr"),"ptr"), "ptr", pIcon.hbmMask)
,w:=pBitmap.bmWidth,h:=pBitmap.bmHeight, h:= h=w*2 ? (h//2,c:=0xffffff,s:=32) : (h,c:=s:=0)
Loop % w {
xi := A_Index - 1
Loop % h {
yi := A_Index - 1 + s
if (DllCall("GetPixel", "ptr", cdc, "uint", xi, "uint", yi) = c) {
if (xo < xi)
xo := xi
if (xs = "" || xs > xi)
xs := xi
if (yo < yi)
yo := yi
if (ys = "" || ys > yi)
ys := yi
}
}
}
DllCall("ReleaseDC", "ptr", 0, "ptr", sdc)
,DllCall("DeleteDC", "ptr", cdc)
,DllCall("DeleteObject", "ptr", hbmo)
,DllCall("DeleteObject", "ptr", picon.hbmMask)
If (y=""){
SysGet,yl,77
SysGet,yr,79
y:=pCursor.y-pIcon.yHotspot+ys+(yo-ys)-s+1
If !(y >= yl && y <= yr)
y:=y<yl ? yl : yr
If (y > yr - 20)
y := yr - 20
}
If (x=""){
SysGet,xr,78
SysGet,xl,76
x:=pCursor.x-pIcon.xHotspot+xs+(xo-xs)+1
If !(x >= xl && x <= xr)
x:=x<xl ? xl : xr
}
}
T.TRACKPOSITION(x,y),T.TRACKACTIVATE(1)
}
TT_Set(T,option:="",OnOff:=1){
static Style:=0x100,NOFADE:=0x20,NoAnimate:=0x10,NOPREFIX:=0x2,AlwaysTip:=0x1,ParseLinks:=0x1000,CloseButton:=0x80,Balloon:=0x40,ClickTrough:=0x20
If (option ~="i)Style|NOFADE|NoAnimate|NOPREFIX|AlwaysTip|ParseLinks|CloseButton|Balloon"){
If ((opt:=DllCall("GetWindowLong","PTR",T.HWND,"UInt",-16) & %option%) && !OnOff) || (!opt && OnOff)
DllCall("SetWindowLong","PTR",T.HWND,"UInt",-16,"UInt",DllCall("GetWindowLong","PTR",T.HWND,"UInt",-16)+(OnOff?(%option%):(-%option%)))
T.Update()
} else If (option="ClickTrough"){
If ((opt:=DllCall("GetWindowLong","PTR",T.HWND,"UInt",-20) & %option%) && !OnOff) || (!opt && OnOff)
DllCall("SetWindowLong","PTR",T.HWND,"UInt",-20,"UInt",DllCall("GetWindowLong","PTR",T.HWND,"UInt",-20)+(OnOff?(%option%):(-%option%)))
T.Update()
} else if !InStr(",__Delete,Push,Pop,InsertAt,Remove,RemoveAt,GetCapacity,SetCapacity,GetAddress,Length,_NewEnum,NewEnum,HasKey,Clone,Count,","," option ",")
MsgBox Invalid option: %option%
}
TT_Font(T, pFont:="") {
static WM_SETFONT := 0x30
italic      := InStr(pFont, "italic")    ?  1    :  0
underline   := InStr(pFont, "underline") ?  1    :  0
strikeout   := InStr(pFont, "strikeout") ?  1    :  0
weight      := InStr(pFont, "bold")      ? 700   : 400
RegExMatch(pFont, "O)(?<=[S|s])(\d{1,2})(?=[ ,])?", height)
if (!height.count())
height := [10]
RegRead,LogPixels,HKEY_LOCAL_MACHINE,SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI, LogPixels
height := -DllCall("MulDiv", "int", Height.1, "int", LogPixels, "int", 72)
RegExMatch(pFont, "O)(?<=,).+", fontFace)
if (fontFace.Value())
fontFace := Trim( fontFace.Value())
else fontFace := "MS Sans Serif"
If (pFont && !InStr(pFont,",") && (italic+underline+strikeout+weight)=400)
fontFace:=pFont
If (T.hFont)
DllCall("DeleteObject","PTR",T.hfont)
T.hFont   := DllCall("CreateFont", "int",  height, "int",  0, "int",  0, "int", 0
,"int",  weight,   "Uint", italic,   "Uint", underline
,"uint", strikeOut, "Uint", nCharSet, "Uint", 0, "Uint", 0, "Uint", 0, "Uint", 0, "str", fontFace,"PTR")
Return DllCall("SendMessage","PTR",T.hwnd,"UInt",WM_SETFONT,"PTR",T.hFont,"PTR",TRUE,"PTR")
}
TTM_ACTIVATE(T,Activate:=0){
static TTM_ACTIVATE := 0x400 + 1
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_ACTIVATE,"PTR",activate,"PTR",0,"PTR")
}
TTM_ADDTOOL(T,pTOOLINFO){
static TTM_ADDTOOL := A_IsUnicode ? 0x432 : 0x404
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_ADDTOOL,"PTR",0,"PTR",pTOOLINFO,"PTR")
}
TTM_ADJUSTRECT(T,action,prect){
static TTM_ADJUSTRECT := 0x41f
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_ADJUSTRECT,"PTR",action,"PTR",prect,"PTR")
}
TTM_DELTOOL(T,pTOOLINFO){
static TTM_DELTOOL := A_IsUnicode ? 0x433 : 0x405
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_DELTOOL,"PTR",0,"PTR",pTOOLINFO,"PTR")
}
TTM_ENUMTOOLS(T,idx,pTOOLINFO){
static TTM_ENUMTOOLS := A_IsUnicode?0x43a:0x40e
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_ENUMTOOLS,"PTR",idx,"PTR",pTOOLINFO,"PTR")
}
TTM_GETBUBBLESIZE(T,pTOOLINFO){
static TTM_GETBUBBLESIZE := 0x41e
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETBUBBLESIZE,"PTR",0,"PTR",pTOOLINFO,"PTR")
}
TTM_GETCURRENTTOOL(T,pTOOLINFO){
static TTM_GETCURRENTTOOL := 0x400 + (A_IsUnicode ? 59 : 15)
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETCURRENTTOOL,"PTR",0,"PTR",pTOOLINFO,"PTR")
}
TTM_GETDELAYTIME(T,whichtime){
static TTM_GETDELAYTIME := 0x400 + 21
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETDELAYTIME,"PTR",whichtime,"PTR",0,"PTR")
}
TTM_GETMARGIN(T,pRECT){
static TTM_GETMARGIN := 0x41b
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETMARGIN,"PTR",0,"PTR",pRECT,"PTR")
}
TTM_GETMAXTIPWIDTH(T,wParam:=0,lParam:=0){
static TTM_GETMAXTIPWIDTH := 0x419
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETMAXTIPWIDTH,"PTR",0,"PTR",0,"PTR")
}
TTM_GETTEXT(T,buffer,pTOOLINFO){
static TTM_GETTEXT := A_IsUnicode?0x438:0x40b
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTEXT,"PTR",buffer,"PTR",pTOOLINFO,"PTR")
}
TTM_GETTIPBKCOLOR(T,wParam:=0,lParam:=0){
static TTM_GETTIPBKCOLOR := 0x416
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTIPBKCOLOR,"PTR",0,"PTR",0,"PTR")
}
TTM_GETTIPTEXTCOLOR(T,wParam:=0,lParam:=0){
static TTM_GETTIPTEXTCOLOR := 0x417
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTIPTEXTCOLOR,"PTR",0,"PTR",0,"PTR")
}
TTM_GETTITLE(T,pTTGETTITLE){
static TTM_GETTITLE := 0x423
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTITLE,"PTR",0,"PTR",pTTGETTITLE,"PTR")
}
TTM_GETTOOLCOUNT(T,wParam:=0,lParam:=0){
static TTM_GETTOOLCOUNT := 0x40d
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTOOLCOUNT,"PTR",0,"PTR",0,"PTR")
}
TTM_GETTOOLINFO(T,pTOOLINFO){
static TTM_GETTOOLINFO := 0x400 + (A_IsUnicode ? 53 : 8)
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_GETTOOLINFO,"PTR",0,"PTR",pTOOLINFO,"PTR")
}
TTM_HITTEST(T,pTTHITTESTINFO){
static TTM_HITTEST := A_IsUnicode?0x437:0x40a
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_HITTEST,"PTR",0,"PTR",pTTHITTESTINFO,"PTR")
}
TTM_NEWTOOLRECT(T,pTOOLINFO:=0){
static TTM_NEWTOOLRECT := 0x400 + (A_IsUnicode ? 52 : 6)
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_NEWTOOLRECT,"PTR",0,"PTR",pTOOLINFO?pTOOLINFO:T.P[],"PTR")
}
TTM_POP(T,wParam:=0,lParam:=0){
static TTM_POP := 0x400 + 28
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_POP,"PTR",0,"PTR",0,"PTR")
}
TTM_POPUP(T,wParam:=0,lParam:=0){
static TTM_POPUP := 0x422
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_POPUP,"PTR",0,"PTR",0,"PTR")
}
TTM_RELAYEVENT(T,wParam:=0,lParam:=0){
static TTM_RELAYEVENT := 0x407
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_RELAYEVENT,"PTR",0,"PTR",0,"PTR")
}
TTM_SETDELAYTIME(T,whichTime:=0,mSec:=-1){
static TTM_SETDELAYTIME := 0x400 + 3
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETDELAYTIME,"PTR",whichTime,"PTR",mSec,"PTR")
}
TTM_SETMARGIN(T,left:=0,top:=0,right:=0,bottom:=0){
static TTM_SETMARGIN := 0x41a
rc:=T.rc,rc.top:=top,rc.left:=left,rc.bottom:=bottom,rc.right:=right
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETMARGIN,"PTR",0,"PTR",rc[],"PTR")
}
TTM_SETMAXTIPWIDTH(T,maxwidth:=-1){
static TTM_SETMAXTIPWIDTH := 0x418
return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETMAXTIPWIDTH,"PTR",0,"PTR",maxwidth,"PTR")
}
TTM_SETTIPBKCOLOR(T,color:=0){
static TTM_SETTIPBKCOLOR := 0x413
,Black:=0x000000,    Green:=0x008000,		Silver:=0xC0C0C0,		Lime:=0x00FF00
,Gray:=0x808080,    	Olive:=0x808000,		White:=0xFFFFFF,   	Yellow:=0xFFFF00
,Maroon:=0x800000,	Navy:=0x000080,		Red:=0xFF0000,    	Blue:=0x0000FF
,Purple:=0x800080,   Teal:=0x008080,		Fuchsia:=0xFF00FF,	Aqua:=0x00FFFF
If InStr(",Black,Green,Silver,Lime,gray,olive,white,yellow,maroon,Navy,Red,Blue,Purple,Teal,Fuchsia,Aqua,","," color ",")
Color:=%color%
Color := (StrLen(Color) < 8 ? "0x" : "") . Color
Color := ((Color&255)<<16)+(((Color>>8)&255)<<8)+(Color>>16)
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETTIPBKCOLOR,"PTR",color,"PTR",0,"PTR")
}
TTM_SETTIPTEXTCOLOR(T,color:=0){
static TTM_SETTIPTEXTCOLOR := 0x414
,Black:=0x000000,    Green:=0x008000,		Silver:=0xC0C0C0,		Lime:=0x00FF00
,Gray:=0x808080,    	Olive:=0x808000,		White:=0xFFFFFF,   	Yellow:=0xFFFF00
,Maroon:=0x800000,	Navy:=0x000080,		Red:=0xFF0000,    	Blue:=0x0000FF
,Purple:=0x800080,   Teal:=0x008080,		Fuchsia:=0xFF00FF,	Aqua:=0x00FFFF
If InStr(",Black,Green,Silver,Lime,gray,olive,white,yellow,maroon,Navy,Red,Blue,Purple,Teal,Fuchsia,Aqua,","," color ",")
Color:=%color%
Color := (StrLen(Color) < 8 ? "0x" : "") . Color
Color := ((Color&255)<<16)+(((Color>>8)&255)<<8)+(Color>>16)
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETTIPTEXTCOLOR,"PTR",color,"PTR",0,"PTR")
}
TTM_SETTITLE(T,title:="",icon:="",Icon_:=1,default:=1){
static TTM_SETTITLE := 0x400 + (A_IsUnicode ? 33 : 32)
If (icon)
If (icon+0="")
If !icon:=TT_GetIcon(icon,Icon_)
icon:=default
T.maintitle := (StrLen(title) < 96) ? title : (Chr(A_IsUnicode ? 8230 : 133) SubStr(title, -96))
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETTITLE,"PTR",icon+0,"PTR",T.GetAddress("maintitle"),"PTR"),T.UPDATE()
}
TTM_SETTOOLINFO(T,pTOOLINFO:=0){
static TTM_SETTOOLINFO := A_IsUnicode?0x436:0x409
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_SETTOOLINFO,"PTR",0,"PTR",pTOOLINFO?pTOOLINFO:T.P[],"PTR")
}
TTM_SETWINDOWTHEME(T,theme:=""){
If !theme
Return DllCall("UxTheme\SetWindowTheme","PTR",T.HWND,"str","","str","")
else Return DllCall("SendMessage","PTR",T.HWND,"UInt",0x200b,"PTR",0,"PTR",&theme,"PTR")
}
TTM_TRACKACTIVATE(T,activate:=0,pTOOLINFO:=0){
Return DllCall("SendMessage","PTR",T.HWND,"UInt",0x411,"PTR",activate,"PTR",(pTOOLINFO)?(pTOOLINFO):(T.P[]),"PTR")
}
TTM_TRACKPOSITION(T,x:=0,y:=0){
Return DllCall("SendMessage","PTR",T.HWND,"UInt",0x412,"PTR",0,"PTR",(x & 0xFFFF)|(y & 0xFFFF)<<16,"PTR")
}
TTM_UPDATE(T){
Return DllCall("SendMessage","PTR",T.HWND,"UInt",0x41D,"PTR",0,"PTR",0,"PTR")
}
TTM_UPDATETIPTEXT(T,pTOOLINFO:=0){
static TTM_UPDATETIPTEXT := A_IsUnicode?0x439:0x40c
Return DllCall("SendMessage","PTR",T.HWND,"UInt",TTM_UPDATETIPTEXT,"PTR",0,"PTR",pTOOLINFO?pTOOLINFO:T.P[],"PTR")
}
TTM_WINDOWFROMPOINT(T,pPOINT){
Return DllCall("SendMessage","PTR",T.HWND,"UInt",0x410,"PTR",0,"PTR",pPOINT,"PTR")
}
_AHKDerefType := "LPTSTR marker,{_AHKVar *var,_AHKFunc *func},BYTE is_function,BYTE param_count,WORD length"
_AHKExprTokenType := "{__int64 value_int64,double value_double,struct{{PTR *object,_AHKDerefType *deref,_AHKVar *var,LPTSTR marker},{LPTSTR buf,size_t marker_length}}},UINT symbol,{_AHKExprTokenType *circuit_token,LPTSTR mem_to_free}"
_AHKArgStruct := "BYTE type,BYTE is_expression,WORD length,LPTSTR text,_AHKDerefType *deref,_AHKExprTokenType *postfix"
_AHKLine := "BYTE ActionType,BYTE Argc,WORD FileIndex,UINT LineNumber,_AHKArgStruct *Arg,PTR *Attribute,*_AHKLine PrevLine,*_AHKLine NextLine,*_AHKLine RelatedLine,*_AHKLine ParentLine"
_AHKLabel := "LPTSTR name,*_AHKLine JumpToLine,*_AHKLabel PrevLabel,*_AHKLabel NextLabel"
_AHKFuncParam := "*_AHKVar var,UShort is_byref,UShort default_type,{default_str,Int64 default_int64,Double default_double}"
If (A_PtrSize = 8)
_AHKRCCallbackFunc := "UINT64 data1,UINT64 data2,PTR stub,UINT_PTR callfuncptr,BYTE actual_param_count,BYTE create_new_thread,event_info,*_AHKFunc func"
else
_AHKRCCallbackFunc := "ULONG data1,ULONG data2,ULONG data3,PTR stub,UINT_PTR callfuncptr,ULONG data4,ULONG data5,BYTE actual_param_count,BYTE create_new_thread,event_info,*_AHKFunc func"
_AHKFunc := "PTR vTable,LPTSTR name,{PTR BIF,*_AHKLine JumpToLine},*_AHKFuncParam Param,Int ParamCount,Int MinParams,*_AHKVar var,*_AHKVar LazyVar,Int VarCount,Int VarCountMax,Int LazyVarCount,Int Instances,*_AHKFunc NextFunc,BYTE DefaultVarType,BYTE IsBuiltIn"
_AHKVar := "{Int64 ContentsInt64,Double ContentsDouble,PTR object},{char *mByteContents,LPTSTR CharContents},{UINT_PTR Length,_AHKVar *AliasFor},{UINT_PTR Capacity,UINT_PTR BIV},BYTE HowAllocated,BYTE Attrib,BYTE IsLocal,BYTE Type,LPTSTR Name"
Class _Struct {
static PTR:=A_PtrSize,UPTR:=A_PtrSize,SHORT:=2,USHORT:=2,INT:=4,UINT:=4,__int64:=8,INT64:=8,UINT64:=8,DOUBLE:=8,FLOAT:=4,CHAR:=1,UCHAR:=1,VOID:=A_PtrSize
,TBYTE:=A_IsUnicode?2:1,TCHAR:=A_IsUnicode?2:1,HALF_PTR:=A_PtrSize=8?4:2,UHALF_PTR:=A_PtrSize=8?4:2,INT32:=4,LONG:=4,LONG32:=4,LONGLONG:=8
,LONG64:=8,USN:=8,HFILE:=4,HRESULT:=4,INT_PTR:=A_PtrSize,LONG_PTR:=A_PtrSize,POINTER_64:=A_PtrSize,POINTER_SIGNED:=A_PtrSize
,BOOL:=4,SSIZE_T:=A_PtrSize,WPARAM:=A_PtrSize,BOOLEAN:=1,BYTE:=1,COLORREF:=4,DWORD:=4,DWORD32:=4,LCID:=4,LCTYPE:=4,LGRPID:=4,LRESULT:=4,PBOOL:=4
,PBOOLEAN:=A_PtrSize,PBYTE:=A_PtrSize,PCHAR:=A_PtrSize,PCSTR:=A_PtrSize,PCTSTR:=A_PtrSize,PCWSTR:=A_PtrSize,PDWORD:=A_PtrSize,PDWORDLONG:=A_PtrSize
,PDWORD_PTR:=A_PtrSize,PDWORD32:=A_PtrSize,PDWORD64:=A_PtrSize,PFLOAT:=A_PtrSize,PHALF_PTR:=A_PtrSize
,UINT32:=4,ULONG:=4,ULONG32:=4,DWORDLONG:=8,DWORD64:=8,ULONGLONG:=8,ULONG64:=8,DWORD_PTR:=A_PtrSize,HACCEL:=A_PtrSize,HANDLE:=A_PtrSize
,HBITMAP:=A_PtrSize,HBRUSH:=A_PtrSize,HCOLORSPACE:=A_PtrSize,HCONV:=A_PtrSize,HCONVLIST:=A_PtrSize,HCURSOR:=A_PtrSize,HDC:=A_PtrSize
,HDDEDATA:=A_PtrSize,HDESK:=A_PtrSize,HDROP:=A_PtrSize,HDWP:=A_PtrSize,HENHMETAFILE:=A_PtrSize,HFONT:=A_PtrSize
static HGDIOBJ:=A_PtrSize,HGLOBAL:=A_PtrSize,HHOOK:=A_PtrSize,HICON:=A_PtrSize,HINSTANCE:=A_PtrSize,HKEY:=A_PtrSize,HKL:=A_PtrSize
,HLOCAL:=A_PtrSize,HMENU:=A_PtrSize,HMETAFILE:=A_PtrSize,HMODULE:=A_PtrSize,HMONITOR:=A_PtrSize,HPALETTE:=A_PtrSize,HPEN:=A_PtrSize
,HRGN:=A_PtrSize,HRSRC:=A_PtrSize,HSZ:=A_PtrSize,HWINSTA:=A_PtrSize,HWND:=A_PtrSize,LPARAM:=A_PtrSize,LPBOOL:=A_PtrSize,LPBYTE:=A_PtrSize
,LPCOLORREF:=A_PtrSize,LPCSTR:=A_PtrSize,LPCTSTR:=A_PtrSize,LPCVOID:=A_PtrSize,LPCWSTR:=A_PtrSize,LPDWORD:=A_PtrSize,LPHANDLE:=A_PtrSize
,LPINT:=A_PtrSize,LPLONG:=A_PtrSize,LPSTR:=A_PtrSize,LPTSTR:=A_PtrSize,LPVOID:=A_PtrSize,LPWORD:=A_PtrSize,LPWSTR:=A_PtrSize,PHANDLE:=A_PtrSize
,PHKEY:=A_PtrSize,PINT:=A_PtrSize,PINT_PTR:=A_PtrSize,PINT32:=A_PtrSize,PINT64:=A_PtrSize,PLCID:=A_PtrSize,PLONG:=A_PtrSize,PLONGLONG:=A_PtrSize
,PLONG_PTR:=A_PtrSize,PLONG32:=A_PtrSize,PLONG64:=A_PtrSize,POINTER_32:=A_PtrSize,POINTER_UNSIGNED:=A_PtrSize,PSHORT:=A_PtrSize,PSIZE_T:=A_PtrSize
,PSSIZE_T:=A_PtrSize,PSTR:=A_PtrSize,PTBYTE:=A_PtrSize,PTCHAR:=A_PtrSize,PTSTR:=A_PtrSize,PUCHAR:=A_PtrSize,PUHALF_PTR:=A_PtrSize,PUINT:=A_PtrSize
,PUINT_PTR:=A_PtrSize,PUINT32:=A_PtrSize,PUINT64:=A_PtrSize,PULONG:=A_PtrSize,PULONGLONG:=A_PtrSize,PULONG_PTR:=A_PtrSize,PULONG32:=A_PtrSize
,PULONG64:=A_PtrSize,PUSHORT:=A_PtrSize,PVOID:=A_PtrSize,PWCHAR:=A_PtrSize,PWORD:=A_PtrSize,PWSTR:=A_PtrSize,SC_HANDLE:=A_PtrSize
,SC_LOCK:=A_PtrSize,SERVICE_STATUS_HANDLE:=A_PtrSize,SIZE_T:=A_PtrSize,UINT_PTR:=A_PtrSize,ULONG_PTR:=A_PtrSize,ATOM:=2,LANGID:=2,WCHAR:=2,WORD:=2,USAGE:=2
static _PTR:="PTR",_UPTR:="UPTR",_SHORT:="Short",_USHORT:="UShort",_INT:="Int",_UINT:="UInt"
,_INT64:="Int64",_UINT64:="UInt64",_DOUBLE:="Double",_FLOAT:="Float",_CHAR:="Char",_UCHAR:="UChar"
,_VOID:="PTR",_TBYTE:=A_IsUnicode?"USHORT":"UCHAR",_TCHAR:=A_IsUnicode?"USHORT":"UCHAR",_HALF_PTR:=A_PtrSize=8?"INT":"SHORT"
,_UHALF_PTR:=A_PtrSize=8?"UINT":"USHORT",_BOOL:="Int",_INT32:="Int",_LONG:="Int",_LONG32:="Int",_LONGLONG:="Int64",_LONG64:="Int64"
,_USN:="Int64",_HFILE:="UInt",_HRESULT:="UInt",_INT_PTR:="PTR",_LONG_PTR:="PTR",_POINTER_64:="PTR",_POINTER_SIGNED:="PTR",_SSIZE_T:="PTR"
,_WPARAM:="PTR",_BOOLEAN:="UCHAR",_BYTE:="UCHAR",_COLORREF:="UInt",_DWORD:="UInt",_DWORD32:="UInt",_LCID:="UInt",_LCTYPE:="UInt"
,_LGRPID:="UInt",_LRESULT:="UInt",_PBOOL:="UPTR",_PBOOLEAN:="UPTR",_PBYTE:="UPTR",_PCHAR:="UPTR",_PCSTR:="UPTR",_PCTSTR:="UPTR"
,_PCWSTR:="UPTR",_PDWORD:="UPTR",_PDWORDLONG:="UPTR",_PDWORD_PTR:="UPTR",_PDWORD32:="UPTR",_PDWORD64:="UPTR",_PFLOAT:="UPTR",___int64:="Int64"
,_PHALF_PTR:="UPTR",_UINT32:="UInt",_ULONG:="UInt",_ULONG32:="UInt",_DWORDLONG:="UInt64",_DWORD64:="UInt64",_ULONGLONG:="UInt64"
,_ULONG64:="UInt64",_DWORD_PTR:="UPTR",_HACCEL:="UPTR",_HANDLE:="UPTR",_HBITMAP:="UPTR",_HBRUSH:="UPTR",_HCOLORSPACE:="UPTR"
,_HCONV:="UPTR",_HCONVLIST:="UPTR",_HCURSOR:="UPTR",_HDC:="UPTR",_HDDEDATA:="UPTR",_HDESK:="UPTR",_HDROP:="UPTR",_HDWP:="UPTR"
static _HENHMETAFILE:="UPTR",_HFONT:="UPTR",_HGDIOBJ:="UPTR",_HGLOBAL:="UPTR",_HHOOK:="UPTR",_HICON:="UPTR",_HINSTANCE:="UPTR",_HKEY:="UPTR"
,_HKL:="UPTR",_HLOCAL:="UPTR",_HMENU:="UPTR",_HMETAFILE:="UPTR",_HMODULE:="UPTR",_HMONITOR:="UPTR",_HPALETTE:="UPTR",_HPEN:="UPTR"
,_HRGN:="UPTR",_HRSRC:="UPTR",_HSZ:="UPTR",_HWINSTA:="UPTR",_HWND:="UPTR",_LPARAM:="UPTR",_LPBOOL:="UPTR",_LPBYTE:="UPTR",_LPCOLORREF:="UPTR"
,_LPCSTR:="UPTR",_LPCTSTR:="UPTR",_LPCVOID:="UPTR",_LPCWSTR:="UPTR",_LPDWORD:="UPTR",_LPHANDLE:="UPTR",_LPINT:="UPTR",_LPLONG:="UPTR"
,_LPSTR:="UPTR",_LPTSTR:="UPTR",_LPVOID:="UPTR",_LPWORD:="UPTR",_LPWSTR:="UPTR",_PHANDLE:="UPTR",_PHKEY:="UPTR",_PINT:="UPTR"
,_PINT_PTR:="UPTR",_PINT32:="UPTR",_PINT64:="UPTR",_PLCID:="UPTR",_PLONG:="UPTR",_PLONGLONG:="UPTR",_PLONG_PTR:="UPTR",_PLONG32:="UPTR"
,_PLONG64:="UPTR",_POINTER_32:="UPTR",_POINTER_UNSIGNED:="UPTR",_PSHORT:="UPTR",_PSIZE_T:="UPTR",_PSSIZE_T:="UPTR",_PSTR:="UPTR"
,_PTBYTE:="UPTR",_PTCHAR:="UPTR",_PTSTR:="UPTR",_PUCHAR:="UPTR",_PUHALF_PTR:="UPTR",_PUINT:="UPTR",_PUINT_PTR:="UPTR",_PUINT32:="UPTR"
,_PUINT64:="UPTR",_PULONG:="UPTR",_PULONGLONG:="UPTR",_PULONG_PTR:="UPTR",_PULONG32:="UPTR",_PULONG64:="UPTR",_PUSHORT:="UPTR"
,_PVOID:="UPTR",_PWCHAR:="UPTR",_PWORD:="UPTR",_PWSTR:="UPTR",_SC_HANDLE:="UPTR",_SC_LOCK:="UPTR",_SERVICE_STATUS_HANDLE:="UPTR"
static _SIZE_T:="UPTR",_UINT_PTR:="UPTR",_ULONG_PTR:="UPTR",_ATOM:="Ushort",_LANGID:="Ushort",_WCHAR:="Ushort",_WORD:="UShort",_USAGE:="UShort"
___InitField(_this,N,offset=" ",encoding=0,AHKType=0,isptr=" ",type=0,arrsize=0,memory=0){
static _prefixes_:={offset:"`b",isptr:"`r",AHKType:"`n",type:"`t",encoding:"`f",memory:"`v",arrsize:" "}
,_testtype_:={offset:"integer",isptr:"integer",AHKType:"string",type:"string",encoding:"string",arrsize:"integer"}
,_default_:={offset:0,isptr:0,AHKType:"UInt",type:"UINT",encoding:"CP0",memory:"",arrsize:1}
for _key_,_value_ in _prefixes_
{
_typevalid_:=0
If (_testtype_[_key_]="Integer"){
If %_key_% is integer
useDefault:=1,_typevalid_:=1
else if !_this.HasKey(_value_ N)
useDefault:=1
} else {
If %_key_% is not integer
useDefault:=1,_typevalid_:=1
else if !_this.HasKey(_value_ N)
useDefault:=1
}
If (useDefault)
If (_key_="encoding")
_this[_value_ N]:=_typevalid_?(InStr(",LPTSTR,LPCTSTR,TCHAR,","," %_key_% ",")?(A_IsUnicode?"UTF-16":"CP0")
:InStr(",LPWSTR,LPCWSTR,WCHAR,","," %_key_% ",")?"UTF-16":"CP0")
:_default_[_key_]
else {
_this[_value_ N]:=_typevalid_?%_key_%:_default_[_key_]
}
}
}
__NEW(_TYPE_,_pointer_=0,_init_=0){
static _base_:={__GET:_Struct.___GET,__SET:_Struct.___SET,__SETPTR:_Struct.___SETPTR,__Clone:_Struct.___Clone,__NEW:_Struct.___NEW
,IsPointer:_Struct.IsPointer,Offset:_Struct.Offset,Type:_Struct.Type,AHKType:_Struct.AHKType,Encoding:_Struct.Encoding
,Capacity:_Struct.Capacity,Alloc:_Struct.Alloc,Size:_Struct.Size,SizeT:_Struct.SizeT,Print:_Struct.Print,ToObj:_Struct.ToObj}
local _,_ArrType_,_ArrName_:="",_ArrSize_,_align_total_,_defobj_,_IsPtr_,_key_,_LF_,_LF_BKP_,_match_,_offset_:=""
,_struct_,_StructSize_,_total_union_size_,_union_,_union_size_,_value_,_mod_,_max_size_,_in_struct_,_struct_align_
If (RegExMatch(_TYPE_,"^[\w\d\._]+$") && !_Struct.HasKey(_TYPE_)){
If InStr(_TYPE_,"."){
Loop,Parse,_TYPE_,.
If A_Index=1
_defobj_:=%A_LoopField%
else _defobj_:=_defobj_[A_LoopField]
_TYPE_:=_defobj_
} else _TYPE_:=%_TYPE_%,_defobj_:=""
} else _defobj_:=""
If (_pointer_ && !IsObject(_pointer_))
this[""] := _pointer_,this["`a"]:=0,this["`a`a"]:=sizeof(_TYPE_)
else
this._SetCapacity("`a",_StructSize_:=sizeof(_TYPE_))
,this[""]:=this._GetAddress("`a")
,DllCall("RtlZeroMemory","UPTR",this[""],"UInt",this["`a`a"]:=_StructSize_)
If InStr(_TYPE_,"`n") {
_struct_:=[]
_union_:=0
Loop,Parse,_TYPE_,`n,`r`t%A_Space%%A_Tab%
{
_LF_:=""
Loop,Parse,A_LoopField,`,`;,`t%A_Space%%A_Tab%
{
If RegExMatch(A_LoopField,"^\s*//")
break
If (A_LoopField){
If (!_LF_ && _ArrType_:=RegExMatch(A_LoopField,"[\w\d_#@]\s+[\w\d_#@]"))
_LF_:=RegExReplace(A_LoopField,"[\w\d_#@]\K\s+.*$")
If Instr(A_LoopField,"{"){
_union_++,_struct_.Insert(_union_,RegExMatch(A_LoopField,"i)^\s*struct\s*\{"))
} else If InStr(A_LoopField,"}")
_offset_.="}"
else {
If _union_
Loop % _union_
_ArrName_.=(_struct_[A_Index]?"struct":"") "{"
_offset_.=(_offset_ ? "," : "") _ArrName_ ((_ArrType_ && A_Index!=1)?(_LF_ " "):"") RegExReplace(A_LoopField,"\s+"," ")
,_ArrName_:="",_union_:=0
}
}
}
}
_TYPE_:=_offset_
}
_offset_:=0
,_union_:=[]
,_struct_:=[]
,_union_size_:=[]
,_struct_align_:=[]
,_total_union_size_:=0
,_align_total_:=0
,_in_struct_:=1
,this["`t"]:=0,this["`r"]:=0
Loop,Parse,_TYPE_,`,`;
{
_in_struct_+=StrLen(A_LoopField)+1
If ("" = _LF_ := trim(A_LoopField,A_Space A_Tab "`n`r"))
continue
_LF_BKP_:=_LF_
_IsPtr_:=0
While (_match_:=RegExMatch(_LF_,"i)^(struct|union)?\s*\{\K"))
_max_size_:=sizeof_maxsize(SubStr(_TYPE_,_in_struct_-StrLen(A_LoopField)-1+(StrLen(_LF_BKP_)-StrLen(_LF_))))
,_union_.Insert(_offset_+=(_mod_:=Mod(_offset_,_max_size_))?Mod(_max_size_-_mod_,_max_size_):0)
,_union_size_.Insert(0)
,_struct_align_.Insert(_align_total_>_max_size_?_align_total_:_max_size_)
,_struct_.Insert(RegExMatch(_LF_,"i)^struct\s*\{")?(1,_align_total_:=0):0)
,_LF_:=SubStr(_LF_,_match_)
StringReplace,_LF_,_LF_,},,A
While % (InStr(_LF_,"*")){
StringReplace,_LF_,_LF_,*
_IsPtr_:=A_Index
}
RegExMatch(_LF_,"^(?<ArrType_>[\w\d\._]+)?\s*(?<ArrName_>[\w\d_]+)?\s*\[?(?<ArrSize_>\d+)?\]?\s*\}*\s*$",_)
If (!_ArrName_ && !_ArrSize_){
If RegExMatch(_TYPE_,"^\**" _ArrType_ "\**$"){
_Struct.___InitField(this,"",0,_ArrType_,_IsPtr_?"PTR":_Struct.HasKey("_" _ArrType_)?_Struct["_" _ArrType_]:"PTR",_IsPtr_,_ArrType_)
this.base:=_base_
If (IsObject(_init_)||IsObject(_pointer_)){
for _key_,_value_ in IsObject(_init_)?_init_:_pointer_
{
If !this["`r"]{
If InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,","," this["`t" _key_] ",")
this.Alloc(_key_,StrLen(_value_)*(InStr(".LPWSTR,LPCWSTR,","," this["`t"] ",")||(InStr(",LPTSTR,LPCTSTR,","," this["`t" _key_] ",")&&A_IsUnicode)?2:1))
if InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,CHAR,TCHAR,WCHAR,","," this["`t" _key_] ",")
this[_key_]:=_value_
else
this[_key_] := _value_
}else if (_value_<>"")
If _value_ is integer
this[_key_][""]:=_value_
}
}
Return this
} else
_ArrName_:=_ArrType_,_ArrType_:="UInt"
}
If InStr(_ArrType_,"."){
Loop,Parse,_ArrType_,.
If A_Index=1
_defobj_:=%A_LoopField%
else _defobj_:=_defobj_[A_LoopField]
}
if (!_IsPtr_ && !_Struct.HasKey(_ArrType_)){
if (sizeof(_defobj_?_defobj_:%_ArrType_%,0,_align_total_) && mod:=Mod(_offset_,_align_total_))
_offset_+=Mod(_align_total_-_mod_,_align_total_)
_Struct.___InitField(this,_ArrName_,_offset_,_ArrType_,0,0,_ArrType_,_ArrSize_)
If (_uix_:=_union_.MaxIndex()) && (_max_size_:=_offset_ + sizeof(_defobj_?_defobj_:%_ArrType_%) - _union_[_uix_])>_union_size_[_uix_]
_union_size_[_uix_]:=_max_size_
_max_size:=0
If (!_uix_||_struct_[_struct_.MaxIndex()])
_offset_+=this[" " _ArrName_]*sizeof(_defobj_?_defobj_:%_ArrType_%)
} else {
If ((_IsPtr_ || _Struct.HasKey(_ArrType_)))
_offset_+=(_mod_:=Mod(_offset_,_max_size_:=_IsPtr_?A_PtrSize:_Struct[_ArrType_]))=0?0:(_IsPtr_?A_PtrSize:_Struct[_ArrType_])-_mod_
,_align_total_:=_max_size_>_align_total_?_max_size_:_align_total_
,_Struct.___InitField(this,_ArrName_,_offset_,_ArrType_,_IsPtr_?"PTR":_Struct.HasKey(_ArrType_)?_Struct["_" _ArrType_]:_ArrType_,_IsPtr_,_ArrType_,_ArrSize_)
If (_uix_:=_union_.MaxIndex()) && (_max_size_:=_offset_ + _Struct[this["`n" _ArrName_]] - _union_[_uix_])>_union_size_[_uix_]
_union_size_[_uix_]:=_max_size_
_max_size_:=0
If (!_uix_||_struct_[_uix_])
_offset_+=_IsPtr_?A_PtrSize:(_Struct.HasKey(_ArrType_)?_Struct[_ArrType_]:%_ArrType_%)*this[" " _ArrName_]
}
While (SubStr(_LF_BKP_,0)="}"){
If (!_uix_:=_union_.MaxIndex()){
MsgBox,0, Incorrect structure, missing opening braket {`nProgram will exit now `n%_TYPE_%
ExitApp
}
if (_uix_>1 && _struct_[_uix_-1]){
if (_mod_:=Mod(_offset_,_struct_align_[_uix_]))
_offset_+=Mod(_struct_align_[_uix_]-_mod_,_struct_align_[_uix_])
} else _offset_:=_union_[_uix_]
if (_struct_[_uix_]&&_struct_align_[_uix_]>_align_total_)
_align_total_ := _struct_align_[_uix_]
_total_union_size_ := _union_size_[_uix_]>_total_union_size_?_union_size_[_uix_]:_total_union_size_
,_union_._Remove(),_struct_._Remove(),_union_size_._Remove(),_struct_align_.Remove(),_LF_BKP_:=SubStr(_LF_BKP_,1,StrLen(_LF_BKP_)-1)
If (_uix_=1){
if (_mod_:=Mod(_total_union_size_,_align_total_))
_total_union_size_ += Mod(_align_total_-_mod_,_align_total_)
_offset_+=_total_union_size_,_total_union_size_:=0
}
}
}
this.base:=_base_
If (IsObject(_init_)||IsObject(_pointer_)){
for _key_,_value_ in IsObject(_init_)?_init_:_pointer_
{
If !this["`r" _key_]{
If InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,","," this["`t" _key_] ",")
this.Alloc(_key_,StrLen(_value_)*(InStr(".LPWSTR,LPCWSTR,","," this["`t"] ",")||(InStr(",LPTSTR,LPCTSTR,","," this["`t" _key_] ",")&&A_IsUnicode)?2:1))
if InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,CHAR,TCHAR,WCHAR,","," this["`t" _key_] ",")
this[_key_]:=_value_
else
this[_key_] := _value_
}else if (_value_<>"")
if _value_ is integer
this[_key_][""]:=_value_
}
}
Return this
}
ToObj(struct:=""){
obj:=[]
for k,v in struct?struct:struct:=this
if (Asc(k)=10)
If IsObject(_VALUE_:=struct[_TYPE_:=SubStr(k,2)])
obj[_TYPE_]:=this.ToObj(_VALUE_)
else obj[_TYPE_]:=_VALUE_
return obj
}
SizeT(_key_=""){
return sizeof(this["`t" _key_])
}
Size(){
return sizeof(this)
}
IsPointer(_key_=""){
return this["`r" _key_]
}
Type(_key_=""){
return this["`t" _key_]
}
AHKType(_key_=""){
return this["`n" _key_]
}
Offset(_key_=""){
return this["`b" _key_]
}
Encoding(_key_=""){
return this["`b" _key_]
}
Capacity(_key_=""){
return this._GetCapacity("`v" _key_)
}
Alloc(_key_="",size="",ptrsize=0){
If _key_ is integer
ptrsize:=size,size:=_key_,_key_:=""
If size is integer
SizeIsInt:=1
If ptrsize {
If (this._SetCapacity("`v" _key_,!SizeIsInt?A_PtrSize+ptrsize:size + (size//A_PtrSize)*ptrsize)="")
MsgBox % "Memory for pointer ." _key_ ". of size " (SizeIsInt?size:A_PtrSize) " could not be set!"
else {
DllCall("RtlZeroMemory","UPTR",this._GetAddress("`v" _key_),"UInt",this._GetCapacity("`v" _key_))
If (this[" " _key_]>1){
ptr:=this[""] + this["`b" _key_]
If (this["`r" _key_] || InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,","," this["`t" _key_] ","))
NumPut(ptrs:=this._GetAddress("`v" _key_),ptr+0,"PTR")
else if _key_
this[_key_,""]:=ptrs:=this._GetAddress("`v" _key_)
else this[""]:=ptr:=this._GetAddress("`v" _key_),ptrs:=this._GetAddress("`v" _key_)+(SizeIsInt?size:A_PtrSize)
} else {
If (this["`r" _key_] || InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR,","," this["`t" _key_] ","))
NumPut(ptr:=this._GetAddress("`v" _key_),this[""] + this["`b" _key_],"PTR")
else this[""]:=ptr:=this._GetAddress("`v" _key_)
ptrs:=ptr+(size?size:A_PtrSize)
}
Loop % SizeIsInt?(size//A_PtrSize):1
NumPut(ptrs+(A_Index-1)*ptrsize,ptr+(A_Index-1)*A_PtrSize,"PTR")
}
} else {
If (this._SetCapacity("`v" _key_,SizeIsInt?size:A_PtrSize)=""){
MsgBox % "Memory for pointer ." _key_ ". of size " (SizeIsInt?size:A_PtrSize) " could not be set!"
} else
NumPut(ptr:=this._GetAddress("`v" _key_),this[""] + this["`b" _key_],"PTR"),DllCall("RtlZeroMemory","UPTR",ptr,"UInt",SizeIsInt?size:A_PtrSize)
}
return ptr
}
___NEW(init*){
this:=this.base
newobj := this.__Clone(1)
If (init.MaxIndex() && !IsObject(init.1))
newobj[""] := init.1
else If (init.MaxIndex()>1 && !IsObject(init.2))
newobj[""] := init.2
else
newobj._SetCapacity("`a",_StructSize_:=sizeof(this))
,newobj[""]:=newobj._GetAddress("`a")
,DllCall("RtlZeroMemory","UPTR",newobj[""],"UInt",_StructSize_)
If (IsObject(init.1)||IsObject(init.2))
for _key_,_value_ in IsObject(init.1)?init.1:init.2
newobj[_key_] := _value_
return newobj
}
___Clone(offset){
static _base_:={__GET:_Struct.___GET,__SET:_Struct.___SET,__SETPTR:_Struct.___SETPTR,__Clone:_Struct.___Clone,__NEW:_Struct.___NEW
,IsPointer:_Struct.IsPointer,Offset:_Struct.Offset,Type:_Struct.Type,AHKType:_Struct.AHKType,Encoding:_Struct.Encoding
,Capacity:_Struct.Capacity,Alloc:_Struct.Alloc,Size:_Struct.Size,SizeT:_Struct.SizeT,Print:_Struct.Print,ToObj:_Struct.ToObj}
If offset=1
return this
newobj:={}
for _key_,_value_ in this
If (_key_!="`a")
newobj[_key_]:=_value_
newobj._SetCapacity("`a",_StructSize_:=sizeof(this))
,newobj[""]:=newobj._GetAddress("`a")
,DllCall("RtlZeroMemory","UPTR",newobj[""],"UInt",_StructSize_)
If this["`r"]{
NumPut(NumGet(this[""],"PTR")+A_PtrSize*(offset-1),newobj[""],"Ptr")
newobj.base:=_base_
} else
newobj.base:=_base_,newobj[]:=this[""]+sizeof(this)*(offset-1)
return newobj
}
___GET(_key_="",p*){
If (_key_="")
Return this[""]
else if !(idx:=p.MaxIndex())
_field_:=_key_,opt:="~"
else {
ObjInsert(p,1,_key_)
opt:=ObjRemove(p),_field_:=_key_:=ObjRemove(p)
for key_,value_ in p
this:=this[value_]
}
If this["`t"]
_key_:=""
If (opt!="~"){
If (opt=""){
If _field_ is integer
return (this["`r"]?NumGet(this[""],"PTR"):this[""])+sizeof(this["`t"])*(_field_-1)
else return this["`r" _key_]?NumGet(this[""]+this["`b" _key_],"PTR"):this[""]+this["`b" _key_]
} else If opt is integer
{
If (_Struct.HasKey("_" this["`t" _key_]) && this[" " _key_]>1) {
If (InStr( ",CHAR,UCHAR,TCHAR,WCHAR," , "," this["`t" _key_] "," )){
Return StrGet(this[""]+this["`b" _key_]+(opt-1)*sizeof(this["`t" _key_]),1,this["`f" _key_])
} else if InStr( ",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t" _key_] "," ){
Return StrGet(NumGet(this[""]+this["`b" _key_]+(opt-1)*A_PtrSize,"PTR"),this["`f" _key_])
} else {
Return NumGet(this[""]+this["`b" _key_]+(opt-1)*sizeof(this["`t" _key_]),this["`n" _key_])
}
} else Return new _Struct(this["`t" _key_],this[""]+this["`b" _key_]+(opt-1)*sizeof(this["`t" _key_]))
} else
return this[_key_][opt]
} else If _field_ is integer
{
If (_key_){
return this.__Clone(_field_)
} else if this["`r"] {
Pointer:=""
Loop % (this["`r"]-1)
pointer.="*"
If pointer
Return new _Struct(pointer this["`t"],NumGet(this[""],"PTR")+A_PtrSize*(_field_-1))
else Return new _Struct(pointer this["`t"],NumGet(this[""],"PTR")+sizeof(this["`t"])*(_field_-1)).1
} else if _Struct.HasKey("_" this["`t"]) {
If (InStr( ",CHAR,UCHAR,TCHAR,WCHAR," , "," this["`t"] "," )){
Return StrGet(this[""]+sizeof(this["`t"])*(_field_-1),1,this["`f"])
} else if InStr(",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t"] "," ){
Return StrGet(NumGet(this[""]+A_PtrSize*(_field_-1),"PTR"),this["`f"])
} else {
Return NumGet(this[""]+sizeof(this["`t"])*(_field_-1),this["`n"])
}
} else {
Return new _Struct(this["`t"],this[""]+sizeof(this["`t"])*(_field_-1))
}
} else If this["`r" _key_] {
Pointer:=""
Loop % (this["`r" _key_]-1)
pointer.="*"
If (_key_=""){
return this[1][_field_]
} else {
Return new _Struct(pointer this["`t" _key_],NumGet(this[""]+this["`b" _key_],"PTR"))
}
} else if _Struct.HasKey("_" this["`t" _key_]) {
If (this[" " _key_]>1)
Return new _Struct(this["`t" _key_],this[""] + this["`b" _key_])
else If (InStr( ",CHAR,UCHAR,TCHAR,WCHAR," , "," this["`t" _key_] "," )){
Return StrGet(this[""]+this["`b" _key_],1,this["`f" _key_])
} else if InStr( ",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t" _key_] "," ){
Return StrGet(NumGet(this[""]+this["`b" _key_],"PTR"),this["`f" _key_])
} else {
Return NumGet(this[""]+this["`b" _key_],this["`n" _key_])
}
} else {
Return new _Struct(this["`t" _key_],this[""]+this["`b" _key_])
}
}
___SET(_key_,p*){
If !(idx:=p.MaxIndex())
return this[""] :=_key_,this._SetCapacity("`a",0)
else if (idx=1)
_value_:=p.1,opt:="~"
else if (idx>1){
ObjInsert(p,1,_key_)
If (p[idx]="")
opt:=ObjRemove(p),_value_:=ObjRemove(p),_key_:=ObjRemove(p)
else _value_:=ObjRemove(p),_key_:=ObjRemove(p),opt:="~"
for key_,value_ in p
this:=this[value_]
}
If this["`t"]
_field_:=_key_,_key_:=""
else _field_:=_key_
If this["`r" _key_] {
If opt is integer
return NumPut(opt,this[""] + this["`b" _key_],"PTR")
else if this.HasKey("`t" _key_) {
Pointer:=""
Loop % (this["`r" _key_]-1)
pointer.="*"
If _key_
Return (new _Struct(pointer this["`t" _key_],NumGet(this[""] + this["`b" _key_],"PTR"))).1:=_value_
else Return (new _Struct(pointer this["`t"],NumGet(this[""],"PTR")))[_field_]:=_value_
} else If _field_ is Integer
if (_key_="")
_this:=this,this:=this.__Clone(_Field_)
If InStr( ",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t" _key_] "," )
StrPut(_value_,NumGet(NumGet(this[""]+this["`b" _key_],"PTR"),"PTR"),this["`f" _key_])
else if InStr( ",TCHAR,CHAR,UCHAR,WCHAR," , "," this["`t" _key_] "," ){
StrPut(_value_,NumGet(this[""]+this["`b" _key_],"PTR"),this["`f" _key_])
} else
NumPut(_value_,NumGet(this[""]+this["`b" _key_],"PTR"),this["`n" _key_])
If _field_ is integer
this:=_this
} else if (RegExMatch(_field_,"^\d+$") && _key_="") {
if InStr( ",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t"] "," ){
StrPut(_value_,NumGet(this[""]+A_PtrSize*(_field_-1),"PTR"),this["`f"])
} else if InStr( ",TCHAR,CHAR,UCHAR,WCHAR," , "," this["`t" _key_] "," ){
StrPut(_value_,this[""] + sizeof(this["`t"])*(_field_-1),this["`f"])
} else
NumPut(_value_,this[""] + sizeof(this["`t"])*(_field_-1),this["`n"])
} else if opt is integer
{
return NumPut(opt,this[""] + this["`b" _key_],"PTR")
} else if InStr( ",LPSTR,LPCSTR,LPTSTR,LPCTSTR,LPWSTR,LPCWSTR," , "," this["`t" _key_] "," ){
StrPut(_value_,NumGet(this[""] + this["`b" _key_],"PTR"),this["`f" _key_])
} else if InStr( ",TCHAR,CHAR,UCHAR,WCHAR," , "," this["`t" _key_] "," ){
StrPut(_value_,this[""] + this["`b" _key_],this["`f" _key_])
} else
NumPut(_value_,this[""]+this["`b" _key_],this["`n" _key_])
Return _value_
}
}
sizeof(_TYPE_,parent_offset=0,ByRef _align_total_=0){
static _types__:="
  (LTrim Join
    ,ATOM:2,LANGID:2,WCHAR:2,WORD:2,PTR:" A_PtrSize ",UPTR:" A_PtrSize ",SHORT:2,USHORT:2,INT:4,UINT:4,INT64:8,UINT64:8,DOUBLE:8,FLOAT:4,CHAR:1,UCHAR:1,__int64:8
    ,TBYTE:" (A_IsUnicode?2:1) ",TCHAR:" (A_IsUnicode?2:1) ",HALF_PTR:" (A_PtrSize=8?4:2) ",UHALF_PTR:" (A_PtrSize=8?4:2) ",INT32:4,LONG:4,LONG32:4,LONGLONG:8
    ,LONG64:8,USN:8,HFILE:4,HRESULT:4,INT_PTR:" A_PtrSize ",LONG_PTR:" A_PtrSize ",POINTER_64:" A_PtrSize ",POINTER_SIGNED:" A_PtrSize "
    ,BOOL:4,SSIZE_T:" A_PtrSize ",WPARAM:" A_PtrSize ",BOOLEAN:1,BYTE:1,COLORREF:4,DWORD:4,DWORD32:4,LCID:4,LCTYPE:4,LGRPID:4,LRESULT:4,PBOOL:" A_PtrSize "
    ,PBOOLEAN:" A_PtrSize ",PBYTE:" A_PtrSize ",PCHAR:" A_PtrSize ",PCSTR:" A_PtrSize ",PCTSTR:" A_PtrSize ",PCWSTR:" A_PtrSize ",PDWORD:" A_PtrSize "
    ,PDWORDLONG:" A_PtrSize ",PDWORD_PTR:" A_PtrSize ",PDWORD32:" A_PtrSize ",PDWORD64:" A_PtrSize ",PFLOAT:" A_PtrSize ",PHALF_PTR:" A_PtrSize "
    ,UINT32:4,ULONG:4,ULONG32:4,DWORDLONG:8,DWORD64:8,ULONGLONG:8,ULONG64:8,DWORD_PTR:" A_PtrSize ",HACCEL:" A_PtrSize ",HANDLE:" A_PtrSize "
     ,HBITMAP:" A_PtrSize ",HBRUSH:" A_PtrSize ",HCOLORSPACE:" A_PtrSize ",HCONV:" A_PtrSize ",HCONVLIST:" A_PtrSize ",HCURSOR:" A_PtrSize ",HDC:" A_PtrSize "
     ,HDDEDATA:" A_PtrSize ",HDESK:" A_PtrSize ",HDROP:" A_PtrSize ",HDWP:" A_PtrSize ",HENHMETAFILE:" A_PtrSize ",HFONT:" A_PtrSize ",USAGE:" 2 "
)"
static _types_:=_types__ "
  (LTrim Join
     ,HGDIOBJ:" A_PtrSize ",HGLOBAL:" A_PtrSize ",HHOOK:" A_PtrSize ",HICON:" A_PtrSize ",HINSTANCE:" A_PtrSize ",HKEY:" A_PtrSize ",HKL:" A_PtrSize "
     ,HLOCAL:" A_PtrSize ",HMENU:" A_PtrSize ",HMETAFILE:" A_PtrSize ",HMODULE:" A_PtrSize ",HMONITOR:" A_PtrSize ",HPALETTE:" A_PtrSize ",HPEN:" A_PtrSize "
     ,HRGN:" A_PtrSize ",HRSRC:" A_PtrSize ",HSZ:" A_PtrSize ",HWINSTA:" A_PtrSize ",HWND:" A_PtrSize ",LPARAM:" A_PtrSize ",LPBOOL:" A_PtrSize ",LPBYTE:" A_PtrSize "
     ,LPCOLORREF:" A_PtrSize ",LPCSTR:" A_PtrSize ",LPCTSTR:" A_PtrSize ",LPCVOID:" A_PtrSize ",LPCWSTR:" A_PtrSize ",LPDWORD:" A_PtrSize ",LPHANDLE:" A_PtrSize "
     ,LPINT:" A_PtrSize ",LPLONG:" A_PtrSize ",LPSTR:" A_PtrSize ",LPTSTR:" A_PtrSize ",LPVOID:" A_PtrSize ",LPWORD:" A_PtrSize ",LPWSTR:" A_PtrSize "
     ,PHANDLE:" A_PtrSize ",PHKEY:" A_PtrSize ",PINT:" A_PtrSize ",PINT_PTR:" A_PtrSize ",PINT32:" A_PtrSize ",PINT64:" A_PtrSize ",PLCID:" A_PtrSize "
     ,PLONG:" A_PtrSize ",PLONGLONG:" A_PtrSize ",PLONG_PTR:" A_PtrSize ",PLONG32:" A_PtrSize ",PLONG64:" A_PtrSize ",POINTER_32:" A_PtrSize "
     ,POINTER_UNSIGNED:" A_PtrSize ",PSHORT:" A_PtrSize ",PSIZE_T:" A_PtrSize ",PSSIZE_T:" A_PtrSize ",PSTR:" A_PtrSize ",PTBYTE:" A_PtrSize "
     ,PTCHAR:" A_PtrSize ",PTSTR:" A_PtrSize ",PUCHAR:" A_PtrSize ",PUHALF_PTR:" A_PtrSize ",PUINT:" A_PtrSize ",PUINT_PTR:" A_PtrSize "
     ,PUINT32:" A_PtrSize ",PUINT64:" A_PtrSize ",PULONG:" A_PtrSize ",PULONGLONG:" A_PtrSize ",PULONG_PTR:" A_PtrSize ",PULONG32:" A_PtrSize "
     ,PULONG64:" A_PtrSize ",PUSHORT:" A_PtrSize ",PVOID:" A_PtrSize ",PWCHAR:" A_PtrSize ",PWORD:" A_PtrSize ",PWSTR:" A_PtrSize ",SC_HANDLE:" A_PtrSize "
     ,SC_LOCK:" A_PtrSize ",SERVICE_STATUS_HANDLE:" A_PtrSize ",SIZE_T:" A_PtrSize ",UINT_PTR:" A_PtrSize ",ULONG_PTR:" A_PtrSize ",VOID:" A_PtrSize "
)"
local _,_ArrName_:="",_ArrType_,_ArrSize_,_defobj_,_idx_,_LF_,_LF_BKP_,_match_,_offset_,_padding_,_struct_
,_total_union_size_,_uix_,_union_,_union_size_,_in_struct_,_mod_,_max_size_,_struct_align_
_offset_:=parent_offset
If IsObject(_TYPE_){
return _TYPE_["`a`a"]
}
If RegExMatch(_TYPE_,"^[\w\d\._]+$"){
If InStr(_types_,"," _TYPE_ ":")
Return SubStr(_types_,InStr(_types_,"," _TYPE_ ":") + 2 + StrLen(_TYPE_),1)
else If InStr(_TYPE_,"."){
Loop,Parse,_TYPE_,.
If A_Index=1
_defobj_:=%A_LoopField%
else _defobj_:=_defobj_[A_LoopField]
Return sizeof(_defobj_,parent_offset)
} else Return sizeof(%_TYPE_%,parent_offset)
} else _defobj_:=""
If InStr(_TYPE_,"`n") {
_offset_:=""
,_struct_:=[]
,_union_:=0
Loop,Parse,_TYPE_,`n,`r`t%A_Space%%A_Tab%
{
_LF_:=""
Loop,Parse,A_LoopField,`,`;,`t%A_Space%%A_Tab%
{
If RegExMatch(A_LoopField,"^\s*//")
break
If (A_LoopField){
If (!_LF_ && _ArrType_:=RegExMatch(A_LoopField,"[\w\d_#@]\s+[\w\d_#@]"))
_LF_:=RegExReplace(A_LoopField,"[\w\d_#@]\K\s+.*$")
If Instr(A_LoopField,"{"){
_union_++,_struct_.Insert(_union_,RegExMatch(A_LoopField,"i)^\s*struct\s*\{"))
} else If InStr(A_LoopField,"}")
_offset_.="}"
else {
If _union_
Loop % _union_
_ArrName_.=(_struct_[A_Index]?"struct":"") "{"
_offset_.=(_offset_ ? "," : "") _ArrName_ ((_ArrType_ && A_Index!=1)?(_LF_ " "):"") RegExReplace(A_LoopField,"\s+"," ")
,_ArrName_:="",_union_:=0
}
}
}
}
_TYPE_:=_offset_
,_offset_:=parent_offset
}
_union_:=[]
,_struct_:=[]
,_union_size_:=[]
,_struct_align_:=[]
,_total_union_size_:=0
,_in_struct_:=1
Loop,Parse,_TYPE_,`,`;
{
_in_struct_+=StrLen(A_LoopField)+1
If ("" = _LF_ := trim(A_LoopField,A_Space A_Tab "`n`r"))
continue
_LF_BKP_:=_LF_
While (_match_:=RegExMatch(_LF_,"i)^(struct|union)?\s*\{\K"))
_max_size_:=sizeof_maxsize(SubStr(_TYPE_,_in_struct_-StrLen(A_LoopField)-1+(StrLen(_LF_BKP_)-StrLen(_LF_))))
,_union_.Insert(_offset_+=(_mod_:=Mod(_offset_,_max_size_))?Mod(_max_size_-_mod_,_max_size_):0)
,_union_size_.Insert(0)
,_struct_align_.Insert(_align_total_>_max_size_?_align_total_:_max_size_)
,_struct_.Insert(RegExMatch(_LF_,"i)^struct\s*\{")?(1,_align_total_:=0):0)
,_LF_:=SubStr(_LF_,_match_)
StringReplace,_LF_,_LF_,},,A
If InStr(_LF_,"*"){
_offset_ += (_mod_:=Mod(_offset_ + A_PtrSize,A_PtrSize)?A_PtrSize-_mod_:0) + A_PtrSize
,_align_total_:=_align_total_<A_PtrSize?A_PtrSize:_align_total_
} else {
RegExMatch(_LF_,"^(?<ArrType_>[\w\d\._#@]+)?\s*(?<ArrName_>[\w\d\._#@]+)?\s*\[?(?<ArrSize_>\d+)?\]?\s*$",_)
If (!_ArrName_ && !_ArrSize_ && !InStr( _types_  ,"," _ArrType_ ":"))
_ArrName_:=_ArrType_,_ArrType_:="UInt"
If InStr(_ArrType_,"."){
Loop,Parse,_ArrType_,.
If A_Index=1
_defobj_:=%A_LoopField%
else _defobj_:=_defobj_[A_LoopField]
}
If (_idx_:=InStr( _types_  ,"," _ArrType_ ":"))
_padding_:=SubStr( _types_  , _idx_+StrLen(_ArrType_)+2 , 1 ),_align_total_:=_align_total_<_padding_?_padding_:_align_total_
else _padding_:= sizeof(_defobj_?_defobj_:%_ArrType_%,0,_align_total_),_max_size_:=sizeof_maxsize(_defobj_?_defobj_:%_ArrType_%)
if (_max_size_){
if (_mod_:=Mod(_offset_,_max_size_))
_offset_ += Mod(_max_size_-_mod_,_max_size_)
} else if _mod_:=Mod(_offset_,_padding_)
_offset_ += Mod(_padding_-_mod_,_padding_)
_offset_ += (_padding_ * (_ArrSize_?_ArrSize_:1))
_max_size_:=0
}
If (_uix_:=_union_.MaxIndex()) && (_max_size_:=_offset_ - _union_[_uix_])>_union_size_[_uix_]
_union_size_[_uix_]:=_max_size_
_max_size_:=0
If (_uix_ && !_struct_[_uix_])
_offset_:=_union_[_uix_]
While (SubStr(_LF_BKP_,0)="}"){
If !(_uix_:=_union_.MaxIndex()){
MsgBox,0, Incorrect structure, missing opening braket {`nProgram will exit now `n%_TYPE_%
ExitApp
}
if (_uix_>1 && _struct_[_uix_-1]){
If (_mod_:=Mod(_offset_,_struct_align_[_uix_]))
_offset_+=Mod(_struct_align_[_uix_]-_mod_,_struct_align_[_uix_])
} else _offset_:=_union_[_uix_]
if (_struct_[_uix_] &&_struct_align_[_uix_]>_align_total_)
_align_total_ := _struct_align_[_uix_]
_total_union_size_ := _union_size_[_uix_]>_total_union_size_?_union_size_[_uix_]:_total_union_size_
,_union_.Remove() ,_struct_.Remove() ,_union_size_.Remove(),_struct_align_.Remove()
,_LF_BKP_:=SubStr(_LF_BKP_,1,StrLen(_LF_BKP_)-1)
If (_uix_=1){
if (_mod_:=Mod(_total_union_size_,_align_total_))
_total_union_size_ += Mod(_align_total_-_mod_,_align_total_)
_offset_+=_total_union_size_,_total_union_size_:=0
}
}
}
_offset_+= Mod(_align_total_ - Mod(_offset_,_align_total_),_align_total_)
Return _offset_
}
sizeof_maxsize(s){
static _types__:="
  (LTrim Join
    ,ATOM:2,LANGID:2,WCHAR:2,WORD:2,PTR:" A_PtrSize ",UPTR:" A_PtrSize ",SHORT:2,USHORT:2,INT:4,UINT:4,INT64:8,UINT64:8,DOUBLE:8,FLOAT:4,CHAR:1,UCHAR:1,__int64:8
    ,TBYTE:" (A_IsUnicode?2:1) ",TCHAR:" (A_IsUnicode?2:1) ",HALF_PTR:" (A_PtrSize=8?4:2) ",UHALF_PTR:" (A_PtrSize=8?4:2) ",INT32:4,LONG:4,LONG32:4,LONGLONG:8
    ,LONG64:8,USN:8,HFILE:4,HRESULT:4,INT_PTR:" A_PtrSize ",LONG_PTR:" A_PtrSize ",POINTER_64:" A_PtrSize ",POINTER_SIGNED:" A_PtrSize "
    ,BOOL:4,SSIZE_T:" A_PtrSize ",WPARAM:" A_PtrSize ",BOOLEAN:1,BYTE:1,COLORREF:4,DWORD:4,DWORD32:4,LCID:4,LCTYPE:4,LGRPID:4,LRESULT:4,PBOOL:" A_PtrSize "
    ,PBOOLEAN:" A_PtrSize ",PBYTE:" A_PtrSize ",PCHAR:" A_PtrSize ",PCSTR:" A_PtrSize ",PCTSTR:" A_PtrSize ",PCWSTR:" A_PtrSize ",PDWORD:" A_PtrSize "
    ,PDWORDLONG:" A_PtrSize ",PDWORD_PTR:" A_PtrSize ",PDWORD32:" A_PtrSize ",PDWORD64:" A_PtrSize ",PFLOAT:" A_PtrSize ",PHALF_PTR:" A_PtrSize "
    ,UINT32:4,ULONG:4,ULONG32:4,DWORDLONG:8,DWORD64:8,ULONGLONG:8,ULONG64:8,DWORD_PTR:" A_PtrSize ",HACCEL:" A_PtrSize ",HANDLE:" A_PtrSize "
     ,HBITMAP:" A_PtrSize ",HBRUSH:" A_PtrSize ",HCOLORSPACE:" A_PtrSize ",HCONV:" A_PtrSize ",HCONVLIST:" A_PtrSize ",HCURSOR:" A_PtrSize ",HDC:" A_PtrSize "
     ,HDDEDATA:" A_PtrSize ",HDESK:" A_PtrSize ",HDROP:" A_PtrSize ",HDWP:" A_PtrSize ",HENHMETAFILE:" A_PtrSize ",HFONT:" A_PtrSize ",USAGE:" 2 "
)"
static _types_:=_types__ "
  (LTrim Join
     ,HGDIOBJ:" A_PtrSize ",HGLOBAL:" A_PtrSize ",HHOOK:" A_PtrSize ",HICON:" A_PtrSize ",HINSTANCE:" A_PtrSize ",HKEY:" A_PtrSize ",HKL:" A_PtrSize "
     ,HLOCAL:" A_PtrSize ",HMENU:" A_PtrSize ",HMETAFILE:" A_PtrSize ",HMODULE:" A_PtrSize ",HMONITOR:" A_PtrSize ",HPALETTE:" A_PtrSize ",HPEN:" A_PtrSize "
     ,HRGN:" A_PtrSize ",HRSRC:" A_PtrSize ",HSZ:" A_PtrSize ",HWINSTA:" A_PtrSize ",HWND:" A_PtrSize ",LPARAM:" A_PtrSize ",LPBOOL:" A_PtrSize ",LPBYTE:" A_PtrSize "
     ,LPCOLORREF:" A_PtrSize ",LPCSTR:" A_PtrSize ",LPCTSTR:" A_PtrSize ",LPCVOID:" A_PtrSize ",LPCWSTR:" A_PtrSize ",LPDWORD:" A_PtrSize ",LPHANDLE:" A_PtrSize "
     ,LPINT:" A_PtrSize ",LPLONG:" A_PtrSize ",LPSTR:" A_PtrSize ",LPTSTR:" A_PtrSize ",LPVOID:" A_PtrSize ",LPWORD:" A_PtrSize ",LPWSTR:" A_PtrSize "
     ,PHANDLE:" A_PtrSize ",PHKEY:" A_PtrSize ",PINT:" A_PtrSize ",PINT_PTR:" A_PtrSize ",PINT32:" A_PtrSize ",PINT64:" A_PtrSize ",PLCID:" A_PtrSize "
     ,PLONG:" A_PtrSize ",PLONGLONG:" A_PtrSize ",PLONG_PTR:" A_PtrSize ",PLONG32:" A_PtrSize ",PLONG64:" A_PtrSize ",POINTER_32:" A_PtrSize "
     ,POINTER_UNSIGNED:" A_PtrSize ",PSHORT:" A_PtrSize ",PSIZE_T:" A_PtrSize ",PSSIZE_T:" A_PtrSize ",PSTR:" A_PtrSize ",PTBYTE:" A_PtrSize "
     ,PTCHAR:" A_PtrSize ",PTSTR:" A_PtrSize ",PUCHAR:" A_PtrSize ",PUHALF_PTR:" A_PtrSize ",PUINT:" A_PtrSize ",PUINT_PTR:" A_PtrSize "
     ,PUINT32:" A_PtrSize ",PUINT64:" A_PtrSize ",PULONG:" A_PtrSize ",PULONGLONG:" A_PtrSize ",PULONG_PTR:" A_PtrSize ",PULONG32:" A_PtrSize "
     ,PULONG64:" A_PtrSize ",PUSHORT:" A_PtrSize ",PVOID:" A_PtrSize ",PWCHAR:" A_PtrSize ",PWORD:" A_PtrSize ",PWSTR:" A_PtrSize ",SC_HANDLE:" A_PtrSize "
     ,SC_LOCK:" A_PtrSize ",SERVICE_STATUS_HANDLE:" A_PtrSize ",SIZE_T:" A_PtrSize ",UINT_PTR:" A_PtrSize ",ULONG_PTR:" A_PtrSize ",VOID:" A_PtrSize "
)"
max:=0,i:=0
s:=trim(s,"`n`r`t ")
If InStr(s,"}"){
Loop,Parse,s
if (A_LoopField="{")
i++
else if (A_LoopField="}"){
if --i<1{
end:=A_Index
break
}
}
if end
s:=SubStr(s,1,end)
}
Loop,Parse,s,`n,`r
{
_struct_:=(i:=InStr(A_LoopField," //"))?SubStr(A_LoopField,1,i):A_LoopField
Loop,Parse,_struct_,`;`,{},%A_Space%%A_Tab%
if A_LoopField&&!InStr(".union.struct.","." A_LoopField ".")
if (!InStr(A_LoopField,A_Tab)&&!InStr(A_LoopField," "))
max:=max<4?4:max
else if (sizeof(A_LoopField,0,size:=0) && max<size)
max:=size
}
return max
}
Struct(Structure,pointer:=0,init:=0)
{
return new _Struct(Structure,pointer,init)
}
Updatechecker:
{
if not (FileExist("C:\Serialmacro\Settings.ini"))
{
FileAppend,, C:\Serialmacro\Settings.ini
Msgbox,4,%A_Space%Serial Macro Updater, %A_Space%Looks like This may be the first time running this application. Do you want to check for an update?
ifmsgbox Yes
{
gosub, Versioncheck
}
Else
{
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
Return
}
}
if (FileExist("C:\Serialmacro\Settings.ini"))
{
IniRead, updatestatus, C:\Serialmacro\Settings.ini, update,lastupdate
IniRead, Updaterate, C:\Serialmacro\Settings.ini, update,updaterate
NumberOfDays := %A_Now%
EnvSub, NumberOfDays, %updatestatus% , Days
If NumberOfDays > %Updaterate% )
{
MsgBox,4,%A_Space%Serial Macro Updater, It has been %NumberOfDays% days since the last update check.`n`n Would you like to check for a new update?`n`n
ifmsgbox Yes
gosub, Versioncheck
else
{
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
}
}
}
Return
}
Versioncheck:
{
Progress,  w200, Updating..., Gathering Information, Serial Macro Updater
Progress, 0
sleep 200
Versioncount = 0
settimer, versiontimeout, 500
gosub, create_checkgui
Progress,  w200, Updating..., Fetching Server Information, Serial Macro Updater
Progress, 15
DllCall("SetParent", "uint",  hwnd, "uint", ParentGUI)
wb.Visible := True
WinSet, Style, -0xC00000, ahk_id %hwnd%
Progress, 25
sleep 200
wb.navigate("https://docs.google.com/document/d/1woiaqcTjqkABrIecRERDAt6nqiEknFWdySqRmie7bCM/edit?usp=sharing")
Progress,  w200, Updating...,Gathering Current Version From Server, Serial Macro Updater
Progress, 50
Sleep, 200
while wb.busy
{
Sleep 100
}
Progress,  w200,Updating..., Comparing Version Information, Serial Macro Updater
Progress, 60
Progress, Off
splashtextoff
Winactivate, Serial version
wingettitle, titleup, A
StringGetPos, pos, Titleup, #, 1
String1 := SubStr(Titleup, pos+2)
Checkversion := SubStr(String1,1,5)
Progress, Off
SplashtextOff
If titleup =
{
Winactivate, Serial Version
wingettitle, titleup, A
StringGetPos, pos, Titleup, #, 1
String1 := SubStr(Titleup, pos+2)
}
If titleup =
{
Progress,  w200,Updating..., Error Occured. Update Not Able To Complete, Serial Macro Updater
Progress, 0
Sleep 3000
Progress, off
SplashTextOff
msgbox, Error getting update. PLease try again later.
}
If Checkversion = %Version_Number%
{
Progress,  w200,Updating..., Macro is Up to date., Serial Macro Updater
Progress, 100
Sleep 3000
Progress, off
SplashTextOff
settimer, versiontimeout, Off
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate
}
If Checkversion > %Version_Number%
{
settimer, versiontimeout, Off
Progress,  w200, Updating..., Update Found. Gathering Update Information, Serial Macro Updater
Progress, 75
Stoploop = 0
endgame = 0
while Stoploop!=10
{
endgame++
If endgame = 15
{
Progress, off
Break
wb.quit
Gui 2:Destroy
Exit
}
++SetTitleMatchMode 2
winactivate, updater
Send ^a
SLEEP 100
ControlGet, Updatetext1,Selected,,,updater
updatetext = Updatetext1
send ^c
updatetext = %clipboard%
SLEEP 500
If updatetext contains Newzz
stoploop = 10
Sleep 1000
}
DOnecount := InStr(updatetext, Donezzz)
Donecount = %donecount% - 1
StringTrimRight,updatetext,updatetext,%donecount%
Loop, Parse, updatetext,`,
{
Whatisnew = %whatisnew%%A_loopField%`n
}
stringreplace, WhatisNew,Whatisnew,newzz,,
stringreplace, WhatisNew,Whatisnew,donezzz,,
Progress, off
SplashTextOff
msgbox,262148,Serial Macro Updater, New update found. Would you like to open the Cat Box site to download the latest version?`n`nThe list below is what has changed.%whatisnew%
IfMsgBox, yes
{
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
Run, https://cat.app.box.com/files/0/f/14285672016/Enter_Serial_Macro
}
Else
Sleep 500
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
return
}
return
}
versiontimeout:
{
Versioncount++
If Versioncount = 60
{
Progress,  w200,Updating..., Error Occured.Cannot connect to server for update check, please check for internet connection., Serial Macro Updater
Progress, 0
Sleep 3000
Progress, Off
SplashTextOff
}
Return
}
create_checkgui:
{
wb := ComObjCreate("InternetExplorer.Application")
Wb.AddressBar := false
wb.MenuBar := false
wb.ToolBar := false
wb.StatusBar := false
wb.Resizable := false
wb.Width := 1000
wb.Height := 500
wb.Top := 0
wb.Left := 0
wb.Silent := true
hwnd := wb.hwnd
Gui, 2:+LastFound
ParentGUI := WinExist()
Gui,2:-SysMenu -ToolWindow -Border
Gui,2:Color, EEAA99
Gui,2:+LastFound
WinSet, TransColor, EEAA99
Gui,2:show, x-10 y-10 W1 H1 , Updater
return
}
$#`:: Listvars