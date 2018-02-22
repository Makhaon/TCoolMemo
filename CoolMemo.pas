(**
* Highlight with TMemo Impossible? try this...
* by Gon Perez-Jimenez May'04
*
* This is a sample how to work with highlighting within TMemo component by
* using interjected class technique.
*
* Of course, this code is still uncompleted but it works fine for my
* purposes, so, hope you can improve it and use it.
*
*)

(*
* Many addtions by Makhaon Software
* AutoComplet from DB
* Multi-level Undo/Redo
* And other. Thanks for Gon Perez-Jimenez for basis code
* (c) 2006 - 2017
*)

unit CoolMemo;

interface

uses
 Winapi.Windows, Winapi.Messages, Vcl.StdCtrls, System.Classes, Vcl.Controls, Vcl.Graphics,
 System.SysUtils, System.Contnrs, Vcl.Forms,
 System.IniFiles, Vcl.Themes, Vcl.Clipbrd, Data.DB, System.Math, System.Types;

type
 TCoolMemo    = class;
 TStrBoolFunc = function(const s: string): boolean of object;

 TPopupListBox = class(TCustomListBox)
 private
  FOldItemIndex: integer;
  function IsKeyPressed(Key: word): boolean;
  procedure WMCancelMode(var Message: TMessage); message WM_CANCELMODE;
 protected
  procedure CreateParams(var Params: TCreateParams); override;
  procedure CreateWnd; override;
  procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: integer); override;
  procedure MouseMove(Shift: TShiftState; X, Y: integer); override;
  procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: integer); override;
 public
  constructor Create(AOwner: TComponent); override;
 end;

 TAutoCompleteLink = class(TDataLink)
 private
  FCoolMemo:      TCoolMemo;
  FDataField:     TField;
  FDataFieldName: string;
  procedure PopulateDataField;
  procedure SetDataFieldName(const Value: string);
  property DataField: TField Read FDataField;
  property DataFieldName: string Read FDataFieldName Write SetDataFieldName;
 protected
  procedure ActiveChanged; override;
 public
  constructor Create(CoolMemo: TCoolMemo);
 end;

 TAutoComplete = class(TPersistent)
 private
  FCoolMemo:      TCoolMemo;
  FPopupListBox:  TPopupListBox;
  FItems:         TStringList;
  FItemIndex:     integer;
  FPrevItemIndex: integer;
  FDropDownCount: integer;
  FDropDownWidth: integer;
  FEnabled:       boolean;
  FVisible:       boolean;
  function FindSelectedItem: boolean;
  function GetItemIndex: integer;
  procedure CloseUp(Apply: boolean);
  function GetPhraseOnPos(const S: string; PhraseDelims: TSysCharSet; Pos: integer): string;
  function GetPhraseOnPositions(const S: string; var PosB, PosE: integer): string;
  procedure ReplacePhrase(const NewString: string);
  procedure SelectItem;
  procedure SetItemIndex(Value: integer);
  procedure SetItems(AItems: TStringList);
  function DoKeyDown(Key: word; Shift: TShiftState): boolean;
  procedure DoKeyPress(Key: char);
 public
  constructor Create(ACoolMemo: TCoolMemo);
  destructor Destroy; override;
  procedure DoCompletion;
  procedure DropDown(AlwaysShow: boolean);
  property Items: TStringList Read FItems Write SetItems;
  property ItemIndex: integer Read GetItemIndex Write SetItemIndex;
  property Visible: boolean Read FVisible Write FVisible;
 published
  property DropDownCount: integer Read FDropDownCount Write FDropDownCount default 6;
  property DropDownWidth: integer Read FDropDownWidth Write FDropDownWidth default 300;
  property Enabled: boolean Read FEnabled Write FEnabled default True;
 end;

 TUndoRedoAction = (uaNone, uaKeyType, uaEnter, uaBackspace, uaDelete, uaCut, uaPaste, uaClear, uaAutoComplete);

 TUndoRedoItem = class
 private
  FAction:   TUndoRedoAction;
  FCaretPos: TPoint;
  FText:     string;
  FSelStart: integer;
  FSelText:  string;
 public
  constructor Create;
 end;

 TUndoRedoList = class(TObjectList)
 private
  FCoolMemo:         TCoolMemo;
  FMaxUndoRedoCount: integer;
  function GetPosBySelStart(SelStart: integer): TPoint;
  procedure SetMaxUndoCount(Value: integer);
  procedure DoUndo;
  procedure DoRedo;
  function AddKeyUndo(const UndoKey: char = #0; FuncKey: boolean = False): TUndoRedoItem;
  function AddClipboardUndo(UndoAction: TUndoRedoAction): TUndoRedoItem;
  function AddAutoCompleteUndo(const AText, ASelText: string): TUndoRedoItem;
  function AddItem(UndoRedoItem: TUndoRedoItem): TUndoRedoItem;
  property MaxUndoRedoCount: integer Read FMaxUndoRedoCount Write SetMaxUndoCount default 1024;
 public
  constructor Create(ACoolMemo: TCoolMemo);
 end;

 TLexemeRec = record
  LexKind: (lkDigit, lkWord);
  Word:    string;
  AddWord: string;
  Color:   TColor;
  Group:   integer;
  Size:    integer;
  Bold:    boolean;
  Italic:  boolean;
 end;

 TLineHighLighter = array of TColor;

 TCoolMemo = class(TCustomMemo)
 private
  FAutoComplete:   TAutoComplete;
  FUndo:           TUndoRedoList;
  FRedo:           TUndoRedoList;
  FIgnoreKeyPress: boolean;
  FShiftKey:       TShiftState;
  FLexemes:        array of TLexemeRec;
  FDataLink:       TAutoCompleteLink;
  FLineHighLihgt:  boolean;
  FLineHighLighter: TLineHighLighter;
  FRightClickMoveCaret: boolean;
  FAddMenu:        boolean;
  FAutoSuggest:    boolean;
  FBackColor:      TColor;
  FOnSpellCheck:   TStrBoolFunc;
  FMinusIsSeparator: boolean;
  function GetDataFieldName: string;
  function GetDataSource: TDataSource;
  function GetMaxUndoCount: integer;
  function IsClipboardKey(Key: char): boolean;
  procedure SetDataFieldName(const Value: string);
  procedure SetDataSource(Value: TDataSource);
  procedure SetFont(Font: Tfont; const word: string; LineNumber: integer);
  procedure SetMaxUndoCount(Value: integer);
  procedure AssignLexem(const FromLexeme: TLexemeRec; ToLexeme: integer);
  procedure ClearLexems;
  function ColStart: integer;
  function CaretCol: integer;
  function Line: integer;
  function TopLine: integer;
  function GetLexKind(const s: string): integer;
  function VisibleLines: integer;
  function CanUndo: boolean;
  function CanRedo: boolean;
  procedure GotoXY(mCol, mLine: integer);
  procedure RefreshDBItems;
  procedure SetBackColor(const Value: TColor);
  function NextWord(var s: string; var PrevWord: string; var SpellOK: boolean): string;
 protected
  procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
  procedure WMSize(var Message: TWMSize); message WM_SIZE;
  procedure WMMove(var Message: TWMMove); message WM_MOVE;
  procedure WMVScroll(var Message: TWMMove); message WM_VSCROLL;
  procedure WMMousewheel(var Message: TWMMove); message WM_MOUSEWHEEL;
  procedure WMCapturechanged(var Message: TWMMove); message WM_CAPTURECHANGED;
  procedure WMEraseBkgnd(var Message: TWMMove); message WM_ERASEBKGND;
  procedure WMKillFocus(var Message: TWMKillFocus); message WM_KILLFOCUS;
  procedure WMCut(var Message); message WM_CUT;
  procedure WMPaste(var Message); message WM_PASTE;
  procedure WMClear(var Message); message WM_CLEAR;
  procedure Change; override;
  procedure KeyDown(var Key: word; Shift: TShiftState); override;
  procedure KeyPress(var Key: char); override;
  procedure KeyUp(var Key: word; Shift: TShiftState); override;
  procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: integer); override;
 public
  constructor Create(AOwner: TComponent); override;
  destructor Destroy; override;
  function LoadLexems(const FileName: string): boolean;
  function GetGroup(const s: string): integer;
  function GetGroupCount(GroupNum: integer): integer;
  function GetWordFromGroup(WordNum, Group: integer; const SelWord: string): string;
  function CaretLine: integer;
  procedure Redo;
  procedure Undo;
  procedure AddLexem(const ALexemeRec: TLexemeRec);
  procedure DelToGroup(Group: integer);
  property LineHighLihgt: boolean Read FLineHighLihgt Write FLineHighLihgt;
  property LineHighLighter: TLineHighLighter Read FLineHighLighter Write FLineHighLighter;
 published
  property AutoSuggest: boolean Read FAutoSuggest Write FAutoSuggest default True;
  property AddMenu: boolean Read FAddMenu Write FAddMenu default False;
  property RightClickMoveCaret: boolean Read FRightClickMoveCaret Write FRightClickMoveCaret;
  property BackColor: TColor Read FBackColor Write SetBackColor default $E1C8FA;
  property DataField: string Read GetDataFieldName Write SetDataFieldName;
  property DataSource: TDataSource Read GetDataSource Write SetDataSource;

  property Align;
  property Alignment;
  property Anchors;
  property AutoComplete: TAutoComplete Read FAutoComplete Write FAutoComplete;
  property BevelEdges;
  property BevelInner;
  property BevelKind default bkNone;
  property BevelOuter;
  property BiDiMode;
  property BorderStyle;
  property Color;
  property Constraints;
  property Ctl3D;
  property DragCursor;
  property DragKind;
  property DragMode;
  property Enabled;
  property Font;
  property HideSelection;
  property ImeMode;
  property ImeName;
  property Lines;
  property MaxLength;
  property OEMConvert;
  property ParentBiDiMode;
  property ParentColor;
  property ParentCtl3D;
  property ParentFont;
  property ParentShowHint;
  property PopupMenu;
  property ReadOnly;
  property ScrollBars;
  property ShowHint;
  property TabOrder;
  property TabStop;
  property MaxUndoCount: integer Read GetMaxUndoCount Write SetMaxUndoCount default 1024;
  property Visible;
  property WantReturns;
  property WantTabs;
  property WordWrap;
  property OnChange;
  property OnClick;
  property OnContextPopup;
  property OnDblClick;
  property OnDragDrop;
  property OnDragOver;
  property OnEndDock;
  property OnEndDrag;
  property OnEnter;
  property OnExit;
  property OnKeyDown;
  property OnKeyPress;
  property OnKeyUp;
  property OnMouseDown;
  property OnMouseMove;
  property OnMouseUp;
  property OnStartDock;
  property OnStartDrag;
  property OnSpellCheck: TStrBoolFunc Read FOnSpellCheck Write FOnSpellCheck;
  property MinusIsSeparator: boolean Read FMinusIsSeparator Write FMinusIsSeparator default True;
 end;

procedure Register;

implementation

const
 LEFT_SHIFT = 1;
 DEFAULT_VALUE = MaxInt;
 AutoCompleteChars = #8 + #46 +
  '_0123456789QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm…÷” ≈Õ√ÿŸ«’⁄‘€¬¿œ–ŒÀƒ∆›ﬂ◊—Ã»“‹¡ﬁ®ÈˆÛÍÂÌ„¯˘Áı˙Ù˚‚‡ÔÓÎ‰Ê˝ˇ˜ÒÏËÚ¸·˛∏';
 Separators: TSysCharSet = [#00, ' ', '-', #13, #10, '.', ',', '/', '\', '#', '"', '''', ':',
  '+', '%', '*', '(', ')', ';', '=', '{', '}', '[', ']', '<', '>'];


 ////////////////////////////////////////////////////////////////////////////////
 // functions for managing keywords and numbers of each line of TMemo ///////////
 ////////////////////////////////////////////////////////////////////////////////

function IsSeparator(Car: char; MinusIsSeparator: boolean): boolean;
begin
 case Car of
  '.', ';', ',', ':', '!', '∑', '"', '''', '^', '+', '*', '/', '\',
  ' ', '`', '[', ']', '(', ')', '{', '}', '?', '%', '=': Result :=
    True;
  '-': Result := MinusIsSeparator;
  else
   Result := False;
 end;
end;
////////////////////////////////////////////////////////////////////////////////

function TCoolMemo.NextWord(var s: string; var PrevWord: string; var SpellOK: boolean): string;
begin
 try
  Result := '';
  PrevWord := '';
  if s = '' then
   Exit;
  while (s <> '') and IsSeparator(s[1], FMinusIsSeparator) do
  begin
   PrevWord := PrevWord + s[1];
   Delete(s, 1, 1);
  end;
  while (s <> '') and not IsSeparator(s[1], FMinusIsSeparator) do
  begin
   Result := Result + s[1];
   Delete(s, 1, 1);
  end;
 finally
  if (Result = '') or (not Assigned(FOnSpellCheck)) then
   SpellOK := True
  else
   SpellOK := FOnSpellCheck(Result);
 end;
end;

function TCoolMemo.VisibleLines: integer;
begin
 Result := Height div (Abs(Self.Font.Height) + 2);
end;

procedure TCoolMemo.GotoXY(mCol, mLine: integer);
begin
 Dec(mLine);
 SelStart  := 0;
 SelLength := 0;
 SelStart  := mCol + Self.Perform(EM_LINEINDEX, mLine, 0);
 SelLength := 0;
 Winapi.Windows.SetFocus(Handle);
end;

function TCoolMemo.TopLine: integer;
begin
 Result := SendMessage(Self.Handle, EM_GETFIRSTVISIBLELINE, 0, 0);
end;

function TCoolMemo.CaretLine: integer;
begin
 Result := SendMessage(Self.Handle, EM_LINEFROMCHAR, WPARAM(Self.SelStart + Self.SelLength), 0);
end;

function TCoolMemo.Line: integer;
begin
 Result := SendMessage(Self.Handle, EM_LINEFROMCHAR, WPARAM(Self.SelStart), 0);
end;

function TCoolMemo.CaretCol: integer;
begin
 Result := Self.SelStart + Self.SelLength - SendMessage(Self.Handle, EM_LINEINDEX, WPARAM(CaretLine), 0);
end;

function TCoolMemo.ColStart: integer;
begin
 Result := Self.SelStart - SendMessage(Self.Handle, EM_LINEINDEX, SendMessage(Self.Handle,
  EM_LINEFROMCHAR, WPARAM(Self.SelStart), 0), 0);
end;

procedure TCoolMemo.WMVScroll(var Message: TWMMove);
begin
 Invalidate;
 inherited;
end;

procedure TCoolMemo.WMSize(var Message: TWMSize);
begin
 Invalidate;
 inherited;
end;

procedure TCoolMemo.WMMove(var Message: TWMMove);
begin
 Invalidate;
 inherited;
end;

procedure TCoolMemo.WMMousewheel(var Message: TWMMove);
begin
 Invalidate;
 inherited;
end;

procedure TCoolMemo.Change;
begin
 Invalidate;
 inherited Change;
end;

procedure TCoolMemo.WMPaint(var Message: TWMPaint);
var
 PS: TPaintStruct;
 DC: HDC;
 Canvas: TCanvas;
 i:  integer;
 X, Y: integer;
 Size, Size2: TSize;
 Max: integer;
 s, Palabra, PrevWord, CurrStr: string;
 TempSelStart, TempSelEnd: integer;
 TempColStart, TempColEnd: integer;
 SpellOK: boolean;
begin
 DC := Message.DC;
 if DC = 0 then
  DC := BeginPaint(Handle, PS);
 Canvas := TCanvas.Create;
 try
  TempSelStart := Line;
  TempSelEnd := CaretLine;
  TempColStart := ColStart;
  TempColEnd := CaretCol;
  try
   Canvas.Handle := DC;
   Canvas.Font.Name := Font.Name;
   Canvas.Font.Size := Font.Size;
   with Canvas do
   begin
    Max := TopLine + VisibleLines;
    if Max > Pred(Lines.Count) then
     Max := Pred(Lines.Count);

    Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
    FillRect(Rect(0, 0, Width, 1));
    Y := 1;
    for i := TopLine to Max do
    begin
     if InRange(i, TempSelStart, TempSelEnd) and (Self.SelLength > 0) then
     begin
      CurrStr := Lines[i];
      if (TempColStart > 0) and (i = TempSelStart) and (TempSelStart = TempSelEnd) then
      begin
       X := LEFT_SHIFT;
       s := Copy(Lines[i], 1, TempColStart);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       TextOut(X, Y, s);
       GetTextExtentPoint32(DC, PChar(s), Length(s), Size);
       Inc(X, Size.cx);
       s := Copy(Lines[i], TempColStart + 1, TempColEnd - TempColStart);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextSelected);
       Brush.Color := clHighlight;
       TextOut(X, Y, s);
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
       GetTextExtentPoint32(DC, PChar(s), Length(s), Size);
       Inc(X, Size.cx);
       s := Copy(Lines[i], TempColEnd + 1, MaxInt);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
       TextOut(X, Y, s);
      end
      else if (TempColStart > 0) and (i = TempSelStart) then
      begin
       X := LEFT_SHIFT;
       s := Copy(Lines[i], 1, TempColStart);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       TextOut(X, Y, s);
       GetTextExtentPoint32(DC, PChar(s), Length(s), Size);
       Inc(X, Size.cx);
       s := Copy(Lines[i], TempColStart + 1, MaxInt);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextSelected);
       Brush.Color := clHighlight;
       TextOut(X, Y, s);
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
      end
      else if (TempColEnd < Length(Lines[i])) and (i = TempSelEnd) then
      begin
       X := LEFT_SHIFT;
       s := Copy(Lines[i], 1, TempColEnd);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextSelected);
       Brush.Color := clHighlight;
       TextOut(X, Y, s);
       GetTextExtentPoint32(DC, PChar(s), Length(s), Size);
       Inc(X, Size.cx);
       s := Copy(Lines[i], TempColEnd + 1, MaxInt);
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
       TextOut(X, Y, s);
      end
      else
      begin
       s := Lines[i];
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextSelected);
       Brush.Color := clHighlight;
       TextOut(LEFT_SHIFT, Y, s);
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
      end;
     end
     else
     begin
      X := LEFT_SHIFT;
      s := Lines[i];
      CurrStr := s;
      //Detecto todas las palabras de esta lÌnea
      Palabra := NextWord(s, PrevWord, SpellOK);
      while Palabra <> '' do
      begin
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       TextOut(X, Y, PrevWord);
       GetTextExtentPoint32(DC, PChar(PrevWord), Length(PrevWord), Size);
       Inc(X, Size.cx);

       SetFont(Font, Palabra, i);
       if not SpellOK then
        Brush.Color := FBackColor;
       TextOut(X, Y, Palabra);
       GetTextExtentPoint32(DC, PChar(Palabra), Length(Palabra), Size);
       Inc(X, Size.cx);

       Palabra := NextWord(s, PrevWord, SpellOK);
       if not SpellOK then
        Brush.Color := FBackColor;
       if (s = '') and (PrevWord <> '') then
       begin
        Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
        TextOut(X, Y, PrevWord);
       end;
       Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
      end;
      if (s = '') and (PrevWord <> '') then
      begin
       Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal);
       TextOut(X, Y, PrevWord);
      end;
     end;
     s := 'W';
     GetTextExtentPoint32(DC, PChar(s), Length(s), Size);
     GetTextExtentPoint32(DC, PChar(CurrStr), Length(CurrStr), Size2);
     Brush.Color:= StyleServices.GetStyleColor(scEditDisabled);
     FillRect(Rect(0, Y, LEFT_SHIFT, Y + Size.cy));
     FillRect(Rect(Size2.cx + LEFT_SHIFT, Y, Width, Y + Size.cy));
     Inc(Y, Size.cy);
    end;
    if Y < Height then
     FillRect(Rect(0, Y, Width, Height));
   end;
  finally
   if Message.DC = 0 then
    EndPaint(Handle, PS);
  end;
 finally
  FreeAndNil(Canvas);
 end;
 inherited;
end;

procedure TCoolMemo.WMCapturechanged(var Message: TWMMove);
begin
 Invalidate;
 inherited;
end;

procedure TCoolMemo.WMEraseBkgnd(var Message: TWMMove);
begin
 Message.Result := 1;
end;

function TCoolMemo.LoadLexems(const FileName: string): boolean;
var
 Groups, Words: integer;
 IniFile: TIniFile;
 i, j: integer;
begin
 Result := False;
 ClearLexems;
 if not FileExists(FileName) then
  Exit;
 IniFile := TIniFile.Create(FileName);
 try
  Groups := IniFile.ReadInteger('Common', 'Groups', -1);
  SetLength(FLexemes, 1);
  with FLexemes[0] do
  begin
   LexKind := lkDigit;
   Word  := '';
   Color := IniFile.ReadInteger('Digits', 'Color', clRed);
   Size  := IniFile.ReadInteger('Digits', 'Size', 0);
   Bold  := IniFile.ReadBool('Digits', 'Bold', False);
   Italic := IniFile.ReadBool('Digits', 'Italic', False);
   Group := 0;
  end;
  for i := 1 to Groups do
  begin
   Words := IniFile.ReadInteger('Group' + IntToStr(i), 'Words', -1);
   for j := 1 to Words do
   begin
    SetLength(FLexemes, Length(FLexemes) + 1);
    with FLexemes[High(FLexemes)] do
    begin
     LexKind := lkWord;
     AddWord := IniFile.ReadString('Group' + IntToStr(i), 'Word' + IntToStr(j), '');
     Word  := AnsiUpperCase(AddWord);
     Color := StringToColor(IniFile.ReadString('Group' + IntToStr(i), 'Color', ColorToString(clBlack)));
     Size  := IniFile.ReadInteger('Group' + IntToStr(i), 'Size', 0);
     Bold  := IniFile.ReadBool('Group' + IntToStr(i), 'Bold', False);
     Italic := IniFile.ReadBool('Group' + IntToStr(i), 'Italic', False);
     Group := i;
    end;
   end;
  end;
 finally
  FreeAndNil(IniFile);
 end;
 Result := True;
end;

procedure TCoolMemo.SetFont(Font: Tfont; const word: string; LineNumber: integer);
var
 TempLexKind: integer;
begin
 if FLineHighLihgt then
 begin
  if LineNumber > High(FLineHighLighter) then
   Font.Color := Self.Font.Color
  else
   Font.Color := FLineHighLighter[LineNumber];
 end
 else
 begin
  TempLexKind := GetLexKind(word);
  if TempLexKind = DEFAULT_VALUE then
   Font.Color := StyleServices.GetStyleFontColor(sfEditBoxTextNormal)
  else
   with FLexemes[TempLexKind] do
    Font.Color := Color;
 end;
end;

function TCoolMemo.GetLexKind(const s: string): integer;
var
 i:  integer;
 s1: string;
begin
 Result := DEFAULT_VALUE;
 if Assigned(FLexemes) then
 begin
  s1 := AnsiUpperCase(s);
  if TryStrToInt(s1, i) then
  begin
   Result := 0;
   exit;
  end;
  for i := 0 to High(FLexemes) do
   if (FLexemes[i].Word = s1) and (FLexemes[i].LexKind <> lkDigit) then
   begin
    Result := i;
    exit;
   end;
 end;
end;

procedure Register;
begin
 RegisterComponents('Makhaon', [TCoolMemo]);
end;

function TCoolMemo.GetGroupCount(GroupNum: integer): integer;
var
 i: integer;
begin
 Result := 0;
 for i := 0 to High(FLexemes) do
  if FLexemes[i].Group = GroupNum then
   Inc(Result);
end;

function TCoolMemo.GetWordFromGroup(WordNum, Group: integer; const SelWord: string): string;
var
 FoundWord, i: integer;
 s1: string;
begin
 Result := '';
 FoundWord := 0;
 s1 := AnsiUpperCase(SelWord);
 for i := 0 to High(FLexemes) do
 begin
  if (FLexemes[i].Group = Group) and (FLexemes[i].Word <> SelWord) then
   Inc(FoundWord);
  if FoundWord = WordNum then
  begin
   Result := (FLexemes[i].AddWord);
   Exit;
  end;
 end;
end;

function TCoolMemo.GetGroup(const s: string): integer;
var
 i:  integer;
 s1: string;
begin
 Result := DEFAULT_VALUE;
 if s = '' then
  Exit;
 s1 := AnsiUpperCase(s);
 for i := 0 to High(FLexemes) do
  if FLexemes[i].Word = s1 then
  begin
   Result := FLexemes[i].Group;
   Break;
  end;
end;

constructor TCoolMemo.Create(AOwner: TComponent);
begin
 inherited Create(AOwner);
 FAutoComplete := TAutoComplete.Create(Self);
 FUndo := TUndoRedoList.Create(Self);
 FRedo := TUndoRedoList.Create(Self);
 FDataLink := TAutoCompleteLink.Create(Self);
 FBackColor := $E1C8FA;
 FRightClickMoveCaret := False;
 FAutoSuggest := True;
 FMinusIsSeparator := True;
end;

destructor TCoolMemo.Destroy;
begin
 FreeAndNil(FDataLink);
 FreeAndNil(FRedo);
 FreeAndNil(FUndo);
 FreeAndNil(FAutoComplete);
 inherited Destroy;
end;

procedure TCoolMemo.KeyPress(var Key: char);
begin
 inherited KeyPress(Key);
 if FIgnoreKeyPress then
 begin
  FIgnoreKeyPress := False;
  Key := #0;
 end
 else
 if (not IsClipboardKey(Key)) then
  if not ((FShiftKey = [ssCtrl]) and (Key = 'Z')) then // Ctrl + Z (Undo)
   FUndo.AddKeyUndo(Key)
  else
  begin
   FUndo.DoUndo;
   Key := #0;
  end;
end;

{ TAutoComplete }

procedure TAutoComplete.CloseUp(Apply: boolean);
begin
 FPrevItemIndex := FItemIndex;
 FItemIndex := ItemIndex;
 FPopupListBox.Visible := False;
 FVisible := False;
 if Apply and (ItemIndex > -1) then
  ReplacePhrase(FItems[ItemIndex]); // Put word/phrase from auto-complete list
end;

constructor TAutoComplete.Create(ACoolMemo: TCoolMemo);
begin
 inherited Create;
 FCoolMemo := ACoolMemo;
 FItems := TStringList.Create;
 FPrevItemIndex := -1;
 FItemIndex := -1;
 FDropDownCount := 6;
 FDropDownWidth := 300;
 FEnabled := True;

 FPopupListBox := TPopupListBox.Create(FCoolMemo);
end;

destructor TAutoComplete.Destroy;
begin
 FreeAndNil(FItems);
 FreeAndNil(FPopupListBox);
 inherited;
end;

procedure TAutoComplete.DoCompletion;
begin
 if FCoolMemo.ReadOnly then
  Exit;
 if FPopupListBox.Visible then
  CloseUp(False);
 DropDown(True);
end;

function TAutoComplete.DoKeyDown(Key: word; Shift: TShiftState): boolean;
begin
 Result := True;
 case Key of
  VK_ESCAPE:
   CloseUp(False);
  VK_RETURN:
   CloseUp(True);
  VK_UP, VK_DOWN, VK_PRIOR, VK_NEXT:
   FPopupListBox.Perform(WM_KEYDOWN, Key, 0);
  else
   Result := False;
 end;
end;

procedure TAutoComplete.DoKeyPress(Key: char);
begin
 if FVisible then
  if Pos(Key, AutoCompleteChars) > 0 then
   SelectItem
  else
   CloseUp(True);
end;

procedure TAutoComplete.DropDown(AlwaysShow: boolean);
var
 ItemsCount: integer;
 ListHeight, ListWidth: integer;
 X, Y, AddedLineIndex: integer;
 R: longint;
 P, OldCursorPos: TPoint;
begin
 CloseUp(False);
 if (not FEnabled) or (FItems.Count = 0) or (not FindSelectedItem and not AlwaysShow) then
  Exit;
 FPopupListBox.Items := FItems;
 FPopupListBox.Style := lbOwnerDrawFixed;
 FPopupListBox.ItemHeight := Round(FCoolMemo.Font.Size * 1.5);
 FVisible  := True;
 ItemIndex := FItemIndex;
 ItemsCount := FItems.Count;

 X := 0;
 Y := 0;
 with FCoolMemo do
 begin
  Lines.BeginUpdate;
  OldCursorPos := GetCaretPos;
  AddedLineIndex := Lines.Add('');
  GotoXY(OldCursorPos.X, OldCursorPos.Y + 1);

  R := Perform(EM_POSFROMCHAR, SelStart + SelLength, 0);
  if R >= 0 then
  begin
   X := LoWord(R);
   Y := HiWord(R);
  end;
  P := Parent.ClientToScreen(Point(Left, Top));
  X := P.X + X + 4;
  Y := P.Y + Y + 4 + Abs(FCoolMemo.Font.Height) + 2;

  Lines.Delete(AddedLineIndex);
  GotoXY(OldCursorPos.X, OldCursorPos.Y + 1);
  Lines.EndUpdate;
 end;

 if ItemsCount > FDropDownCount then
  ItemsCount := FDropDownCount;
 ListHeight := FPopupListBox.ItemHeight * ItemsCount + 2;
 ListWidth := FDropDownWidth;
 if ListWidth = 0 then
  ListWidth := FCoolMemo.Width + 2 * GetSystemMetrics(SM_CXBORDER);
 FPopupListBox.Left := X;
 FPopupListBox.Top := Y;
 FPopupListBox.Width := ListWidth;
 FPopupListBox.Height := ListHeight;
 if Y + FPopupListBox.Height > Screen.Height then
  Y := Y - FPopupListBox.Height - (Abs(FCoolMemo.Font.Height) + 2 + 4);
 SetWindowPos(FPopupListBox.Handle, HWND_TOP, X, Y, 0, 0,
  SWP_NOSIZE or SWP_NOACTIVATE or SWP_SHOWWINDOW);
 FPopupListBox.Visible := True;
end;

function TAutoComplete.FindSelectedItem: boolean;
var
 s: string;
 i, j, k: integer;
 PhraseDelims: TSysCharSet;
begin
 for j := 1 to 3 do
 begin
  case j of
   1: PhraseDelims := ['.']; // separeted by dot
   2: PhraseDelims := [#13, #10]; // separated by line break
   else
    PhraseDelims := [' ']; // separated by space
  end;

  with FCoolMemo do
   if Lines.Count > 0 then
    s := TrimLeft(GetPhraseOnPos(Lines[CaretLine], PhraseDelims, CaretCol))
   else
    s := '';

  k := -1;
  if s <> '' then
  begin
   i := 0;
   while i <= FItems.Count - 1 do
   begin
    if ANSIStrLIComp(PChar(FItems[i]), PChar(s), Length(s)) = 0 then
    begin
     k := i;
     Break;
    end;
    Inc(i);
   end;
   if k > -1 then
    Break;
  end;
 end;

 ItemIndex := k;
 Result := ItemIndex > -1;
end;

function TAutoComplete.GetItemIndex: integer;
begin
 Result := FItemIndex;
 if FVisible then
  Result := FPopupListBox.ItemIndex;
end;

function TAutoComplete.GetPhraseOnPos(const S: string; PhraseDelims: TSysCharSet; Pos: integer): string;
var
 I, Count: integer;
begin
 I := Pos;
 while (I > 0) and not (CharInSet(S[I], PhraseDelims)) do
  Dec(I);
 Count := Pos - I;
 if (I > 0) and (CharInSet(S[I], PhraseDelims)) then
  Inc(I);
 Result := Copy(S, I, Count);
end;

function TAutoComplete.GetPhraseOnPositions(const S: string; var PosB, PosE: integer): string;
var
 N, Len: integer;
 AutoCompleteStr: string;
 PhraseDelims: TSysCharSet;
 Found:  boolean;
begin
 if ItemIndex > -1 then
 begin
  AutoCompleteStr := FItems[ItemIndex];
  // Find begin index from cursor position to the start of line
  if FPrevItemIndex > -1 then
  begin
   for N := 1 to 2 do
   begin
    case N of
     1: PhraseDelims := ['.']; // separeted by dot
     2: PhraseDelims := [#13, #10]; // separated by line break
    end;
    Found := False;
    while (PosB > 0) and not (CharInSet(S[PosB], PhraseDelims)) do
     Dec(PosB);
    if (PosB > 0) and (CharInSet(S[PosB], PhraseDelims)) then
     Found := True;
    Inc(PosB);
    while (PosB <= PosE) and (CharInSet(S[PosB], [' '])) do
     Inc(PosB);
    if Found then
     Break;
   end;
  end
  else
  begin
   while (PosB > 0) and not (CharInSet(S[PosB], Separators)) do
    Dec(PosB);
   Inc(PosB);
  end;
  // Find end index from cursor position to the end of the current word
  Len := Length(S);
  if Len > 0 then
  begin
   while (PosE < Len) and not (CharInSet(S[PosE], Separators)) do
    Inc(PosE);
   if CharInSet(S[PosE], Separators) then
    Dec(PosE);
  end;
  Result := AutoCompleteStr;
 end
 else
  Result := '';
end;

procedure TAutoComplete.ReplacePhrase(const NewString: string);
var
 S, Phrase: string;
 LineIndex, PosB, PosE: integer;
begin
 with FCoolMemo do
 begin
  Lines.BeginUpdate;
  LineIndex := CaretLine;
  S := Lines[LineIndex];
  PosB := CaretCol;
  PosE := PosB;
  Phrase := GetPhraseOnPositions(S, PosB, PosE);
  if Phrase <> '' then
  begin
   FUndo.AddAutoCompleteUndo(Phrase, Copy(S, PosB, PosE - PosB + 1));
   Delete(S, PosB, PosE - PosB + 1);
   Insert(NewString, S, PosB);
   Lines[LineIndex] := S;
   GotoXY(PosB + Length(Phrase) - 1, LineIndex + 1);
  end;
  Lines.EndUpdate;
 end;
end;

procedure TAutoComplete.SelectItem;
begin
 FindSelectedItem;
 if (not Visible and (ItemIndex = -1)) or (FItems.Count = 0) then
  CloseUp(False);
end;

procedure TAutoComplete.SetItemIndex(Value: integer);
begin
 FItemIndex := Value;
 if FVisible then
  FPopupListBox.ItemIndex := FItemIndex;
end;

procedure TAutoComplete.SetItems(AItems: TStringList);
begin
 FItems.Assign(AItems);
end;

procedure TCoolMemo.KeyDown(var Key: word; Shift: TShiftState);
var
 AutoKey: word;
begin
 inherited KeyDown(Key, Shift);

 AutoKey := Key;
 FShiftKey := Shift;

 if FAutoComplete.Visible then
 begin
  if (Key in [VK_UP, VK_DOWN]) then
   Key := 0; // Hold cursor at the same position
  if FAutoComplete.DoKeyDown(AutoKey, Shift) then
  begin
   FIgnoreKeyPress := True;
   Exit;
  end;
 end;

 if (Key = VK_SPACE) and (Shift = [ssCtrl]) then
 begin
  FIgnoreKeyPress := True;
  FAutoComplete.DropDown(True);
 end;

 if Key = VK_DELETE then
  FUndo.AddKeyUndo(Chr(Key), True);
end;

{ TAutoCompleteList }

procedure TCoolMemo.WMKillFocus(var Message: TWMKillFocus);
begin
 inherited;
 if csFocusing in ControlState then
  Exit;
 if FAutoComplete.Visible then
  FAutoComplete.CloseUp(False);
end;

procedure TCoolMemo.KeyUp(var Key: word; Shift: TShiftState);
var
 CharKey: char;
begin
 inherited KeyUp(Key, Shift);

 CharKey := Chr(Key);
 if Pos(CharKey, AutoCompleteChars) > 0 then
  FAutoComplete.DoKeyPress(CharKey);
end;

procedure TCoolMemo.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
 p, Len: integer;
begin
 inherited;
 if FAutoComplete.Visible then
  FAutoComplete.CloseUp(False);
 SetFocus;
 if (Button = mbRight) and FRightClickMoveCaret then
 begin
  Len := SendMessage(Handle, WM_GETTEXTLENGTH, 0, 0);
  if Len < 65536 then
   p := LoWord(SendMessage(Handle, EM_CHARFROMPOS, 0, MakeLParam(X, Y)))
  else
   p := SendMessage(Handle, EM_CHARFROMPOS, 0, MakeLParam(X, Y));
  SendMessage(Handle, EM_SETSEL, WPARAM(p), LPARAM(p));
 end;
end;

{ TPopupListbox }

constructor TPopupListBox.Create(AOwner: TComponent);
begin
 inherited Create(AOwner);
 Parent := AOwner as TCoolMemo;
 IntegralHeight := True;
 ItemHeight := 13;
 Left := -1000;
 ShowHint := True;
 Visible := False;

 FOldItemIndex := -1;
end;

procedure TPopupListbox.CreateParams(var Params: TCreateParams);
begin
 inherited CreateParams(Params);
 with Params do
 begin
  Style := Style or WS_BORDER;
  ExStyle := WS_EX_TOOLWINDOW or WS_EX_TOPMOST;
  WindowClass.Style := CS_SAVEBITS;
 end;
end;

procedure TPopupListbox.CreateWnd;
begin
 inherited CreateWnd;
 Winapi.Windows.SetParent(Handle, 0);
 CallWindowProc(DefWndProc, Handle, WM_SETFOCUS, 0, 0);
end;

function TPopupListbox.IsKeyPressed(Key: word): boolean;
begin
 Result := GetKeyState(Key) and $8000 = $8000;
end;

procedure TPopupListbox.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
 CurItemIndex: integer;
begin
 inherited;
 MouseCapture := True;
 CurItemIndex := ItemAtPos(Point(X, Y), True);
 if CurItemIndex > -1 then
  ItemIndex := CurItemIndex;
end;

procedure TPopupListbox.MouseMove(Shift: TShiftState; X, Y: integer);
var
 CurItemIndex: integer;
begin
 CurItemIndex := ItemAtPos(Point(X, Y), True);

 if (CurItemIndex >= 0) and (CurItemIndex <> FOldItemIndex) then
 begin
  FOldItemIndex := CurItemIndex;
  Application.ProcessMessages;
  Application.CancelHint;
  Hint := '';
  if Canvas.TextWidth(Items[CurItemIndex]) > Width - 4 then
   Hint := Items[CurItemIndex];
 end;

 if IsKeyPressed(VK_LBUTTON) and (CurItemIndex > -1) then
  ItemIndex := CurItemIndex;
end;

procedure TPopupListbox.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: integer);
begin
 MouseCapture := False;
 (Owner as TCoolMemo).AutoComplete.CloseUp((Button = mbLeft) and PtInRect(ClientRect, Point(X, Y)));
end;

procedure TPopupListbox.WMCancelMode(var Message: TMessage);
begin
 (Owner as TCoolMemo).AutoComplete.CloseUp(False);
end;

{ TUndoRedoItem }

constructor TUndoRedoItem.Create;
begin
 inherited Create;
 FAction := uaNone;
 FCaretPos := Point(0, 0);
 FSelStart := -1;
end;

{ TUndoRedoList }

function TUndoRedoList.AddItem(UndoRedoItem: TUndoRedoItem): TUndoRedoItem;
begin
 Result := TUndoRedoItem.Create;
 with Result do
 begin
  FAction := UndoRedoItem.FAction;
  FCaretPos := UndoRedoItem.FCaretPos;
  FText := UndoRedoItem.FText;
  FSelStart := UndoRedoItem.FSelStart;
  FSelText := UndoRedoItem.FSelText;
 end;
 if Count >= FMaxUndoRedoCount then
  Delete(0);
 Add(Result);
end;

function TUndoRedoList.AddKeyUndo(const UndoKey: char = #0; FuncKey: boolean = False): TUndoRedoItem;
begin
 Result := TUndoRedoItem.Create;
 with FCoolMemo do
 begin
  if SelLength > 0 then
  begin
   Result.FSelStart := SelStart;
   Result.FSelText  := Copy(Text, SelStart + 1, SelLength);
  end;
  Result.FCaretPos := CaretPos;
  if Ord(UndoKey) = VK_RETURN then
   Result.FAction := uaEnter
  else if Ord(UndoKey) = VK_BACK then
  begin
   Result.FAction := uaBackspace;
   if Result.FSelText = '' then
    if Result.FCaretPos.X > 0 then // Is first char in the string
     Result.FText := Lines[Result.FCaretPos.Y][Result.FCaretPos.X]
    else
    begin
     Result.FText := Chr(VK_RETURN);
     if Result.FCaretPos.Y > 0 then
     begin
      Dec(Result.FCaretPos.Y);
      Result.FCaretPos.X := Length(Lines[Result.FCaretPos.Y]);
     end;
    end;
  end
  else if (Ord(UndoKey) = VK_DELETE) and FuncKey then
  begin
   Result.FAction := uaDelete;
   if Result.FSelText = '' then
    if (Result.FCaretPos.X < Length(Lines[Result.FCaretPos.Y])) then
     Result.FText := Lines[Result.FCaretPos.Y][Result.FCaretPos.X + 1]
    else
     Result.FText := Chr(VK_RETURN);
  end
  else
  begin
   Result.FAction := uaKeyType;
   Result.FText := UndoKey;
   if Result.FSelText <> '' then
    Result.FCaretPos := GetPosBySelStart(SelStart);
  end;
  if ((Ord(UndoKey) = VK_BACK) and (Result.FCaretPos.X = 0) and (Result.FCaretPos.Y = 0) or
   (Ord(UndoKey) = VK_DELETE) and (Result.FCaretPos.X = Length(Lines[Result.FCaretPos.Y])) and
   (Result.FCaretPos.Y = Lines.Count) and (Result.FSelText = '')) then
   FreeAndNil(Result)
  else
  begin
   if Count >= FMaxUndoRedoCount then
    Delete(0);
   Add(Result);
   if FRedo.Count > 0 then
    FRedo.Clear;
  end;
 end;
end;

function TUndoRedoList.AddClipboardUndo(UndoAction: TUndoRedoAction): TUndoRedoItem;
begin
 Result := TUndoRedoItem.Create;
 with FCoolMemo do
 begin
  Result.FAction := UndoAction;
  Result.FCaretPos := CaretPos;
  if UndoAction = uaPaste then
   Result.FText := Clipboard.AsText;
  Result.FSelStart := SelStart;
  if SelLength > 0 then
   Result.FSelText := Copy(Text, SelStart + 1, SelLength);
  if Count >= FMaxUndoRedoCount then
   Delete(0);
  Add(Result);
  if FRedo.Count > 0 then
   FRedo.Clear;
 end;
end;

function TUndoRedoList.AddAutoCompleteUndo(const AText, ASelText: string): TUndoRedoItem;
begin
 Result := TUndoRedoItem.Create;
 with FCoolMemo, Result do
 begin
  FAction := uaAutoComplete;
  FCaretPos := CaretPos;
  FText := AText;
  FSelStart := FSelStart - Length(ASelText);
  FSelText := ASelText;
  if Count >= FMaxUndoRedoCount then
   Delete(0);
  Add(Result);
  if FRedo.Count > 0 then
   FRedo.Clear;
 end;
end;

constructor TUndoRedoList.Create(ACoolMemo: TCoolMemo);
begin
 inherited Create;
 FCoolMemo := ACoolMemo;
 FMaxUndoRedoCount := 1024;
end;

procedure TUndoRedoList.DoRedo;
var
 UndoRedo: TUndoRedoItem;
 S: string;
 NewPos: TPoint;
begin
 with FCoolMemo do
  if CanRedo then
  begin
   Lines.BeginUpdate;
   UndoRedo := TUndoRedoItem(Self.Last);
   if UndoRedo.FSelText <> '' then
   begin
    S := Text;
    System.Delete(S, UndoRedo.FSelStart + 1, Length(UndoRedo.FSelText));
    Text := S;
    NewPos := GetPosBySelStart(UndoRedo.FSelStart);
   end
   else
    NewPos := UndoRedo.FCaretPos;
   S := Lines[NewPos.Y];
   case UndoRedo.FAction of
    uaKeyType:
    begin
     System.Insert(UndoRedo.FText, S, NewPos.X + 1);
     Inc(NewPos.X);
    end;
    uaEnter:
    begin
     Lines.Insert(NewPos.Y + 1, Copy(S, NewPos.X + 1, Length(S)));
     System.Delete(S, NewPos.X + 1, Length(S));
     NewPos.X := 0;
    end;
    uaBackspace:
     if UndoRedo.FSelText = '' then
      if Ord(UndoRedo.FText[1]) <> VK_RETURN then
      begin
       System.Delete(S, NewPos.X, Length(UndoRedo.FText));
       Dec(NewPos.X);
      end
      else
      begin
       S := S + Lines[NewPos.Y + 1];
       Lines.Delete(NewPos.Y + 1);
      end;
    uaDelete:
     if UndoRedo.FSelText = '' then
      if Ord(UndoRedo.FText[1]) <> VK_RETURN then
       System.Delete(S, NewPos.X + 1, Length(UndoRedo.FText))
      else
      begin
       S := S + Lines[NewPos.Y + 1];
       Lines.Delete(NewPos.Y + 1);
      end;
   end;
   Lines[NewPos.Y] := S;
   if UndoRedo.FAction = uaEnter then
    Inc(NewPos.Y);
   if (UndoRedo.FAction = uaPaste) or (UndoRedo.FAction = uaAutoComplete) then
   begin
    S := Text;
    System.Insert(UndoRedo.FText, S, UndoRedo.FSelStart + 1);
    Text := S;
    NewPos := GetPosBySelStart(UndoRedo.FSelStart + Length(UndoRedo.FText));
   end;
   GotoXY(NewPos.X, NewPos.Y + 1);
   FUndo.AddItem(UndoRedo);
   Delete(Count - 1);
   Lines.EndUpdate;
  end;
end;

procedure TUndoRedoList.DoUndo;
var
 UndoRedo: TUndoRedoItem;
 S: string;
 NewPos: TPoint;
begin
 with FCoolMemo do
  if CanUndo then
  begin
   Lines.BeginUpdate;
   UndoRedo := TUndoRedoItem(Self.Last);
   S := Lines[UndoRedo.FCaretPos.Y];
   NewPos := UndoRedo.FCaretPos;
   case UndoRedo.FAction of
    uaKeyType:
     System.Delete(S, UndoRedo.FCaretPos.X + 1, Length(UndoRedo.FText));
    uaEnter:
    begin
     S := S + Lines[UndoRedo.FCaretPos.Y + 1];
     Lines.Delete(UndoRedo.FCaretPos.Y + 1);
    end;
    uaBackspace:
     if UndoRedo.FSelText = '' then
      if Ord(UndoRedo.FText[1]) <> VK_RETURN then
       System.Insert(UndoRedo.FText, S, UndoRedo.FCaretPos.X)
      else
      begin
       Lines.Insert(UndoRedo.FCaretPos.Y + 1,
        Copy(S, UndoRedo.FCaretPos.X + 1, Length(S)));
       System.Delete(S, UndoRedo.FCaretPos.X + 1, Length(S));
       NewPos.X := 0;
       Inc(NewPos.Y);
      end;
    uaDelete:
     if UndoRedo.FSelText = '' then
      if Ord(UndoRedo.FText[1]) <> VK_RETURN then
       System.Insert(UndoRedo.FText, S, UndoRedo.FCaretPos.X + 1)
      else
      begin
       Lines.Insert(UndoRedo.FCaretPos.Y + 1, Copy(S, UndoRedo.FCaretPos.X + 1, Length(S)));
       System.Delete(S, UndoRedo.FCaretPos.X + 1, Length(S));
      end;
    uaPaste:
     if UndoRedo.FSelText = '' then
     begin
      S := Text;
      System.Delete(S, UndoRedo.FSelStart + 1, Length(UndoRedo.FText));
     end;
   end;
   if (UndoRedo.FAction <> uaPaste) then
    Lines[UndoRedo.FCaretPos.Y] := S
   else
   if UndoRedo.FSelText = '' then
    Text := S;
   if UndoRedo.FSelText <> '' then
   begin
    S := Text;
    if (UndoRedo.FAction = uaPaste) or (UndoRedo.FAction = uaAutoComplete) then
     System.Delete(S, UndoRedo.FSelStart + 1, Length(UndoRedo.FText));
    System.Insert(UndoRedo.FSelText, S, UndoRedo.FSelStart + 1);
    Text := S;
    // Set new cursor position
    if (UndoRedo.FAction = uaKeyType) and (UndoRedo.FSelText <> '') then
     NewPos := GetPosBySelStart(SelStart + Length(UndoRedo.FSelText) + 1);
   end;
   GotoXY(NewPos.X, NewPos.Y + 1);
   FRedo.AddItem(UndoRedo);
   Delete(Count - 1);
   Lines.EndUpdate;
  end;
end;

procedure TCoolMemo.Undo;
begin
 FUndo.DoUndo;
end;

function TUndoRedoList.GetPosBySelStart(SelStart: integer): TPoint;
begin
 Result.Y := SendMessage(FCoolMemo.Handle, EM_LINEFROMCHAR, WPARAM(SelStart), 0);
 Result.X := SelStart - SendMessage(FCoolMemo.Handle, EM_LINEINDEX, WPARAM(Result.Y), 0);
end;

procedure TUndoRedoList.SetMaxUndoCount(Value: integer);
begin
 while (FMaxUndoRedoCount > Value) and (FMaxUndoRedoCount > 0) do
 begin
  Delete(0);
  Dec(FMaxUndoRedoCount);
 end;
 FMaxUndoRedoCount := Value;
end;

procedure TCoolMemo.Redo;
begin
 FRedo.DoRedo;
end;

procedure TCoolMemo.WMCut(var Message);
begin
 FUndo.AddClipboardUndo(uaCut);
 inherited;
end;

procedure TCoolMemo.WMPaste(var Message);
begin
 FUndo.AddClipboardUndo(uaPaste);
 inherited;
end;

procedure TCoolMemo.WMClear(var Message);
begin
 FUndo.AddClipboardUndo(uaClear);
 inherited;
end;

function TCoolMemo.GetMaxUndoCount: integer;
begin
 Result := FUndo.MaxUndoRedoCount;
end;

procedure TCoolMemo.SetMaxUndoCount(Value: integer);
begin
 FUndo.MaxUndoRedoCount := Value;
 FRedo.MaxUndoRedoCount := Value;
end;

function TCoolMemo.CanUndo: boolean;
begin
 Result := not ReadOnly and Enabled and (FUndo.Count > 0);
end;

function TCoolMemo.CanRedo: boolean;
begin
 Result := not ReadOnly and Enabled and (FRedo.Count > 0);
end;

function TCoolMemo.GetDataFieldName: string;
begin
 Result := FDataLink.DataFieldName;
end;

function TCoolMemo.GetDataSource: TDataSource;
begin
 Result := FDataLink.DataSource;
end;

procedure TCoolMemo.SetBackColor(const Value: TColor);
begin
 if FBackColor <> Value then
 begin
  FBackColor := Value;
  if not (csDesigning in ComponentState) then
   Refresh;
 end;
end;

procedure TCoolMemo.SetDataFieldName(const Value: string);
begin
 FDataLink.DataFieldName := Value;
end;

procedure TCoolMemo.SetDataSource(Value: TDataSource);
begin
 if FDataLink.DataSource <> Value then
 begin
  FDataLink.DataSource := Value;
  if not (csLoading in ComponentState) then
   FDataLink.DataFieldName := '';
 end;
end;

procedure TCoolMemo.RefreshDBItems;
var
 BookMk: TBookmark;
begin
 with FDataLink do
  if (DataField <> nil) and (DataSource <> nil) and (DataSource.DataSet <> nil) and (DataSource.DataSet.Active) then
  begin
   FAutoComplete.Items.Clear;
   with DataSource.DataSet do
   begin
    BookMk := GetBookmark;
    try
     DisableControls;
     First;
     while not Eof do
     begin
      FAutoComplete.Items.Add(FieldByName(FDataLink.DataFieldName).AsString);
      Next;
     end;
     GotoBookmark(BookMk);
     EnableControls;
    finally
     FreeBookmark(BookMk);
    end;
   end;
  end;
end;

function TCoolMemo.IsClipboardKey(Key: char): boolean;
begin
 Result := (FShiftKey = [ssCtrl]) and ((Key = #22) or (Key = #24));
end;

{ TAutoCompleteLink }

procedure TAutoCompleteLink.ActiveChanged;
begin
 if not (csDesigning in FCoolMemo.ComponentState) then
  if Active then
   PopulateDataField
  else
  begin
   FDataField := nil;
   FCoolMemo.AutoComplete.Items.Clear;
  end;
end;

constructor TAutoCompleteLink.Create(CoolMemo: TCoolMemo);
begin
 inherited Create;
 FCoolMemo  := CoolMemo;
 FDataField := nil;
end;

procedure TAutoCompleteLink.PopulateDataField;
begin
 if (DataSource <> nil) and (DataSource.DataSet <> nil) and (DataSource.DataSet.Active) and
  (FDataFieldName <> '') then
 begin
  FDataField := DataSource.DataSet.FieldByName(FDataFieldName);
  FCoolMemo.RefreshDBItems;
 end;
end;

procedure TAutoCompleteLink.SetDataFieldName(const Value: string);
begin
 if FDataFieldName <> Value then
 begin
  FDataFieldName := Value;
  PopulateDataField;
 end;
end;

procedure TCoolMemo.AssignLexem(const FromLexeme: TLexemeRec; ToLexeme: integer);
begin
 with FLexemes[ToLexeme] do
 begin
  LexKind := FromLexeme.LexKind;
  Word  := FromLexeme.Word;
  Color := FromLexeme.Color;
  Size  := FromLexeme.Size;
  Bold  := FromLexeme.Bold;
  Italic := FromLexeme.Italic;
  Group := FromLexeme.Group;
 end;
end;

procedure TCoolMemo.AddLexem(const ALexemeRec: TLexemeRec);
begin
 SetLength(FLexemes, Length(FLexemes) + 1);
 AssignLexem(ALexemeRec, High(FLexemes));
end;

procedure TCoolMemo.DelToGroup(Group: integer);
var
 i: integer;
begin
 for i := 0 to High(FLexemes) do
  if FLexemes[i].Group = Group then
  begin
   SetLength(FLexemes, i);
   Break;
  end;
end;

procedure TCoolMemo.ClearLexems;
begin
 Finalize(FLexemes, 0);
end;

end.