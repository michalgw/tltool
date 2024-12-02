program tltool;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Classes, SysUtils, CustApp,
  { you can add units after this }
  LazUTF8, FileInfo, winpeimagereader, LazFileUtils, ActiveX, ComObj,
  DOM, XMLRead, XMLWrite, XPath;

type

  { TTypeLibTool }

  TTypeLibTool = class(TCustomApplication)
  private
    type TOperation = (oGenLibManifest, oUpdateVersion);
  protected
    Operation: TOperation;
    InputType: (itDllExe, itTlb);
    InputFile: String;
    LibVersion: String;
    OutFile: String;
    DALibName: String;
    procedure DoRun; override;
    procedure GetVersion;
    procedure MakeLibManifest;
    procedure UpdateVersionInManifest;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteHelp; virtual;
  end;

{ TManifestTool }

procedure TTypeLibTool.DoRun;
var
  ErrorMsg: String;
  S: String;
begin
  // quick check parameters
  ErrorMsg := CheckOptions('o:i:m:v:l:h', 'operation: input: manifest: version: libname: help');
  if ErrorMsg <> '' then begin
    ShowException(Exception.Create(ErrorMsg));
    Terminate;
    Exit;
  end;

  // parse parameters
  if HasOption('h', 'help') then begin
    WriteHelp;
    Terminate;
    Exit;
  end;

  S := GetOptionValue('o', 'operation');
  case UpperCase(S) of
    'GEN': Operation := oGenLibManifest;
    'UPD': Operation := oUpdateVersion;
    else begin
      ShowException(Exception.Create('Invalid operation: ' + S));
      Terminate;
      Exit;
    end;
  end;

  InputFile := GetOptionValue('i', 'input');
  if (InputFile = '') and (Operation = oGenLibManifest) then
  begin
    ShowException(Exception.Create('Input file not specified'));
    Terminate;
    Exit;
  end;
  if SameText(ExtractFileExt(InputFile), '.tlb') then
    InputType := itTlb
  else
    InputType := itDllExe;

  OutFile := GetOptionValue('m', 'manifest');
  if (OutFile = '') then
    if Operation = oGenLibManifest then
      OutFile := InputFile + '.manifest'
    else
    begin
      ShowException(Exception.Create('No manifest file specified'));
      Terminate;
      Exit;
    end;

  LibVersion := GetOptionValue('v', 'version');
  DALibName := GetOptionValue('l', 'libname');

  if (Operation = oUpdateVersion) and ((InputFile = '') or (InputType = itTlb)) then
  begin
    if LibVersion = '' then
    begin
      ShowException(Exception.Create('No version given'));
      Terminate;
      Exit;
    end;
    if (DALibName = '') and (InputFile = '') then
    begin
      ShowException(Exception.Create('Library name missing'));
      Terminate;
      Exit;
    end;
  end;

  { add your program here }
  case Operation of
    oGenLibManifest: MakeLibManifest;
    oUpdateVersion: UpdateVersionInManifest;
  end;

  // stop program loop
  Terminate;
end;

procedure TTypeLibTool.GetVersion;
var
  FVer: TFileVersionInfo;
begin
  if LibVersion <> '' then
    Exit;
  if InputType = itDllExe then
  begin
    FVer := TFileVersionInfo.Create(nil);
    try
      FVer.FileName := InputFile;
      FVer.ReadFileInfo;
      LibVersion := FVer.VersionStrings.Values['FileVersion'];
    finally
      FVer.Free;
    end;
  end;
  if LibVersion = '' then
    LibVersion := '1.0.0.0';
end;

procedure TTypeLibTool.MakeLibManifest;
var
  TLib: ITypeLib;
  LAtt: lpTLIBATTR;
  I: Integer;
  TyInfTy: TYPEKIND;
  TyInf: ITypeInfo;
  TyAtt: LPTYPEATTR;
  LibName: WideString;
  BstrName: WideString;
  XML: TStringList = nil;
  LFName: String;
begin
  GetVersion;
  XML := TStringList.Create;
  try
    OleCheck(LoadTypeLib(PWideChar(WideString(InputFile)), TLib));
    OleCheck(TLib.GetLibAttr(LAtt));
    OleCheck(TLib.GetDocumentation(-1, @LibName, nil, nil, nil));
    with XML do
    begin
      Add('<?xml version="1.0" encoding="utf-8" standalone="yes"?>');
      Add('<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">');
      Add('  <assemblyIdentity name="' + ExtractFileNameOnly(InputFile) + '" type="win32" version="' + LibVersion + '" />');
      LFName := ExtractFileName(InputFile);
      if InputType = itTlb then
        LFName := ChangeFileExt(LFName, '.dll');
      Add('  <file name="' + LFName + '">');
      Add('    <typelib tlbid="' + GUIDToString(LAtt^.GUID) + '" version="' + IntToStr(LAtt^.wMajorVerNum) + '.' + IntToStr(LAtt^.wMinorVerNum) + '" flags="HASDISKIMAGE" helpdir="" />');
      for I := 0 to TLib.GetTypeInfoCount - 1 do
      begin
        TLib.GetTypeInfoType(I, TyInfTy);
        if TyInfTy = TKIND_COCLASS then
        begin
          OleCheck(TLib.GetTypeInfo(I, TyInf));
          OleCheck(TyInf.GetTypeAttr(TyAtt));
          OleCheck(TyInf.GetDocumentation(DISPID_UNKNOWN, @BstrName, nil, nil, nil));
          Add('    <comClass progid="' + LibName + '.' + BstrName + '" clsid="' + GUIDToString(TyAtt^.GUID) + '"  threadingModel="Apartment" />');
          TyInf.ReleaseTypeAttr(TyAtt);
        end;
      end;
      Add('  </file>');
      for I := 0 to TLib.GetTypeInfoCount - 1 do
      begin
        TLib.GetTypeInfoType(I, TyInfTy);
        if (TyInfTy = TKIND_INTERFACE) or (TyInfTy = TKIND_DISPATCH) then
        begin
          OleCheck(TLib.GetTypeInfo(I, TyInf));
          OleCheck(TyInf.GetTypeAttr(TyAtt));
          OleCheck(TyInf.GetDocumentation(DISPID_UNKNOWN, @BstrName, nil, nil, nil));
          Add('  <comInterfaceExternalProxyStub name="' + BstrName + '" iid="' + GUIDToString(TyAtt^.GUID) + '" proxyStubClsid32="{00020424-0000-0000-C000-000000000046}" baseInterface="{00000000-0000-0000-C000-000000000046}" tlbid="' + GUIDToString(LAtt^.GUID) + '" />');
          TyInf.ReleaseTypeAttr(TyAtt);
        end;
      end;
      Add('</assembly>');
    end;
    TLib.ReleaseTLibAttr(LAtt);
    XML.SaveToFile(OutFile);
  finally
    if Assigned(XML) then
      XML.Free;
  end;
end;

procedure TTypeLibTool.UpdateVersionInManifest;
var
  X: TXMLDocument;
  XR: TXPathNSResolver;
  PR: TXPathVariable = nil;
  N: TDOMElement;
begin
  GetVersion;
  if DALibName = '' then
    DALibName := ExtractFileNameOnly(InputFile);
  ReadXMLFile(X, OutFile);
  XR := TXPathNSResolver.Create(X.DocumentElement);
  PR := EvaluateXPathExpression('/assembly/dependency/dependentAssembly/assemblyIdentity[@name=''' + DALibName + ''']',
    X.DocumentElement, XR);
  if Assigned(PR) and Assigned(PR.AsNodeSet) and (PR.AsNodeSet.Count > 0) then
  begin
    N := TDOMElement(TDOMNode(PR.AsNodeSet.Items[0]));
    N.SetAttribute('version', LibVersion);
    WriteXMLFile(X, OutFile);
  end;
  if Assigned(PR) then
    PR.Free;
  XR.Free;
  X.Free;
end;

constructor TTypeLibTool.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  StopOnException := True;
end;

destructor TTypeLibTool.Destroy;
begin
  inherited Destroy;
end;

procedure TTypeLibTool.WriteHelp;
begin
  WriteLn('Type library tool');
  WriteLn('Usage: ', ExtractFileNameOnly(ExeName), ' -o <gen|upd> -i <input file> -m <manifest file> -v <version> -l <lib name>');
  WriteLn('  -o <gen,upd>    - Type of operation');
  WriteLn('          gen     - Generating a manifest for a library');
  WriteLn('          upd     - Update library version in existing manifest file');
  WriteLn('  -i <input file> - Type library file (tlb/dll/ocx)');
  WriteLn('  -m <manifest file> - Output manifest file or manifest file to be modified');
  WriteLn('  -v <version>    - Specifying library version');
  WriteLn('  -l <lib name>   - Library name');
  WriteLn;
  WriteLn('Example:');
  WriteLn(ExtractFileNameOnly(ExeName), ' -o gen -i comlib.dll -m comlib.dll.manifest');
  WriteLn('  Create COM type library manifest, type library and version from "comlib.dll" file, generated manifest saved to "comlib.dll.manifest".');
  WriteLn;
  WriteLn(ExtractFileNameOnly(ExeName), ' -o upd -m app.exe.manifest -v 2.3.0.0 -l comlib');
  WriteLn('  Update library versions in "/assembly/dependency/dependentAssembly/assemblyIdentity[@name=comlib]" node of manifest file.');
end;

var
  Application: TTypeLibTool;
begin
  Application := TTypeLibTool.Create(nil);
  Application.Title := 'Type library tool';
  Application.Run;
  Application.Free;
end.

