{-----------------------------------------------------------------------------
 Unit Name: gtcXdom31DomUnit
 Author:    Tor Helland
 Purpose:   IDom... interface wrapper for OpenXml Xdom 3.1
 History:
-----------------------------------------------------------------------------}
unit gtcXdom31DomUnit;

interface
uses Classes,
  Variants,
  SysUtils,
  ComObj,
  xmldom,
  Xdom_3_1;

const
  sXdom31Xml = 'Open XML 3.1';                    { Do not localize }

type

{ IOXDOMNodeRef }

  Iox31DOMNodeRef = interface
    ['{4D898FD5-1F65-44E9-9E27-A28026311F94}']
    function GetNativeNode: TdomNode;
  end;

{ Tox31DOMInterface }

  Tox31DOMInterface = class(TInterfacedObject)
  public
    function SafeCallException(ExceptObject: TObject; ExceptAddr: Pointer): HRESULT; override;
  end;

{ Tox31DOMImplementation }


   PParseErrorInfo = ^TParseErrorInfo;
   TParseErrorInfo = record
     errorCode: Integer;
     errorCodeStr: string;
     url: WideString;
     reason: WideString;
     srcText: WideString;
     line: Integer;
     linePos: Integer;
     filePos: Integer;
   end;

  Tox31DOMDocument = class;

  Tox31DOMImplementation = class(Tox31DOMInterface, IDOMImplementation)
  private
    FNativeDOMImpl: TDomImplementation;
    FParser            : TXmlToDomParser;
    FBuilder           : TXmlDomBuilder;
    FReader            : TXmlStandardDomReader;
    FXpath             : TXPathExpression;
    FParseError: PParseErrorInfo;
    function GetNativeDOMImpl: TdomImplementation;
  protected
    { IDOMImplementation }
    function hasFeature(const feature, version: DOMString): WordBool;
    function createDocumentType(const qualifiedName, publicId,
      systemId: DOMString): IDOMDocumentType; safecall;
    function createDocument(const namespaceURI, qualifiedName: DOMString;
      doctype: IDOMDocumentType): IDOMDocument; safecall;
  public
    constructor Create;
    destructor Destroy; override;
    { Parsing Helpers for IDOMPersist }
    procedure FreeDocument(var Doc: TdomDocument);
    procedure InitParserAgent;
    function loadFromStream(const stream: TStream; const WrapperDoc: Tox31DOMDocument;
      var ParseError: TParseErrorInfo): WordBool;
    function loadxml(const Value: DOMString; const WrapperDoc: Tox31DOMDocument;
      var ParseError: TParseErrorInfo): WordBool;
    procedure ParseErrorHandler(sender: TObject; error: TdomError);
    property NativeDOMImpl: TdomImplementation read GetNativeDOMImpl;
  end;

{ Tox31DOMNode }

  Tox31DOMNodeClass = class of Tox31DOMNode;

  Tox31DOMNode = class(Tox31DOMInterface, Iox31DOMNodeRef,
    IDOMNode, IDOMNodeEx, IDOMNodeSelect)
  private
    FNativeNode: TdomNode;
    FWrapperDocument: Tox31DOMDocument;
    FChildNodes: IDOMNodeList;
    FAttributes: IDOMNamedNodeMap;
    FOwnerDocument: IDOMDocument;
  protected
    function AllocParser: TDomToXmlParser; // Must be freed by calling routine.

    { Iox31DOMNodeRef }
    function GetNativeNode: TdomNode;
    { IDOMNode }
    function get_nodeName: DOMString; virtual; safecall;
    function get_nodeValue: DOMString; virtual; safecall;
    procedure set_nodeValue(value: DOMString); virtual; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_nodeType: DOMNodeType; safecall;
    function get_parentNode: IDOMNode; safecall;
    function get_childNodes: IDOMNodeList; virtual; safecall;
    function get_firstChild: IDOMNode; safecall;
    function get_lastChild: IDOMNode; safecall;
    function get_previousSibling: IDOMNode; safecall;
    function get_nextSibling: IDOMNode; safecall;
    function get_attributes: IDOMNamedNodeMap; safecall;
    function get_ownerDocument: IDOMDocument; safecall;
    function get_namespaceURI: DOMString; virtual; safecall;
    function get_prefix: DOMString; virtual; safecall;
    function get_localName: DOMString; virtual; safecall;
    function insertBefore(const newChild, refChild: IDOMNode): IDOMNode; safecall;
    function replaceChild(const newChild, oldChild: IDOMNode): IDOMNode; safecall;
    function removeChild(const childNode: IDOMNode): IDOMNode; safecall;
    function appendChild(const newChild: IDOMNode): IDOMNode; safecall;
    function hasChildNodes: WordBool; virtual; safecall;
    function cloneNode(deep: WordBool): IDOMNode; safecall;
    procedure normalize; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function supports(const feature, version: DOMString): WordBool;
    { IDOMNodeEx }
    function get_text: DOMString; safecall;
    function get_xml: DOMString; safecall;
    procedure set_text(const Value: DOMString); safecall;
    procedure transformNode(const stylesheet: IDOMNode; var output: WideString); overload; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure transformNode(const stylesheet: IDOMNode; const output: IDOMDocument); overload; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    { IDOMNodeSelect }
    function selectNode(const nodePath: WideString): IDOMNode; safecall;
    function selectNodes(const nodePath: WideString): IDOMNodeList; safecall;
  public
    constructor Create(ANativeNode: TdomNode; AWrapperDocument: Tox31DOMDocument); virtual;
    destructor Destroy; override;
    property NativeNode: TdomNode read FNativeNode;
    property WrapperDocument: Tox31DOMDocument read FWrapperDocument;
  end;

{ Tox31DOMNodeList }

  Tox31DOMNodeList = class(Tox31DOMInterface, IDOMNodeList)
  private
     FNativeNodeList: TdomNodeList;
     FNativeXpathNodeSet: TdomXPathCustomResult;
     FWrapperOwnerNode: Tox31DOMNode;
  protected
    { IDOMNodeList }
    function get_item(index: Integer): IDOMNode; safecall;
    function get_length: Integer; safecall;
  public
    constructor Create(ANativeNodeList: TdomNodeList; AWrapperOwnerNode: Tox31DOMNode); overload;
    constructor Create(AnXpath: TXpathExpression; AWrapperOwnerNode: Tox31DOMNode); overload;
    destructor Destroy; override;
    property NativeNodeList: TdomNodeList read FNativeNodeList;
  end;

{ Tox31DOMNamedNodeMap }

  Tox31DOMNamedNodeMap = class(Tox31DOMInterface, IDOMNamedNodeMap)
  private
    FNativeNamedNodeMap: TdomNamedNodeMap;
    FWrapperOwnerNode: Tox31DOMNode;
    procedure CheckNamespaceAware;
  protected
    { IDOMNamedNodeMap }
    function get_item(index: Integer): IDOMNode; safecall;
    function get_length: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function getNamedItem(const name: DOMString): IDOMNode; safecall;
    function setNamedItem(const arg: IDOMNode): IDOMNode; safecall;
    function removeNamedItem(const name: DOMString): IDOMNode; safecall;
    function getNamedItemNS(const namespaceURI, localName: DOMString): IDOMNode; safecall;
    function setNamedItemNS(const arg: IDOMNode): IDOMNode; safecall;
    function removeNamedItemNS(const namespaceURI, localName: DOMString): IDOMNode; safecall;
  public
    constructor Create(ANativeNamedNodeMap: TdomNamedNodeMap; AWrapperOwnerNode: Tox31DOMNode);
    property NativeNamedNodeMap: TdomNamedNodeMap read FNativeNamedNodeMap;
  end;

{ Tox31DOMCharacterData }

  Tox31DOMCharacterData = class(Tox31DOMNode, IDOMCharacterData)
  private
    function GetNativeCharacterData: TdomCharacterData;
  protected
    { IDOMCharacterData }
    function get_data: DOMString; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_data(const data: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_length: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function substringData(offset, count: Integer): DOMString; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure appendData(const data: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure insertData(offset: Integer; const data: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure deleteData(offset, count: Integer); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure replaceData(offset, count: Integer; const data: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
  public
    property NativeCharacterData: TdomCharacterData read GetNativeCharacterData;
  end;

{ Tox31DOMAttr }

  Tox31DOMAttr = class(Tox31DOMNode, IDOMAttr)
  private
    function GetNativeAttribute: TdomAttr;
  protected
    { Property Get/Set }
    function get_name: DOMString; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_specified: WordBool; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_value: DOMString; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_value(const attributeValue: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_ownerElement: IDOMElement; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    { Properties }
    property name: DOMString read get_name;
    property specified: WordBool read get_specified;
    property value: DOMString read get_value write set_value;
    property ownerElement: IDOMElement read get_ownerElement;
  public
    property NativeAttribute: TdomAttr read GetNativeAttribute;
  end;

{ Tox31DOMElement }

  Tox31DOMElement = class(Tox31DOMNode, IDOMElement)
  private
    function GetNativeElement: TdomElement;
    procedure CheckNamespaceAware;
  protected
    { IDOMElement }
    function get_tagName: DOMString; safecall;
    function getAttribute(const name: DOMString): DOMString; safecall;
    procedure setAttribute(const name, value: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure removeAttribute(const name: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function getAttributeNode(const name: DOMString): IDOMAttr; safecall;
    function setAttributeNode(const newAttr: IDOMAttr): IDOMAttr; safecall;
    function removeAttributeNode(const oldAttr: IDOMAttr): IDOMAttr; safecall;
    function getElementsByTagName(const name: DOMString): IDOMNodeList;
      safecall;
    function getAttributeNS(const namespaceURI, localName: DOMString):
      DOMString; safecall;
    procedure setAttributeNS(const namespaceURI, qualifiedName,
      value: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure removeAttributeNS(const namespaceURI, localName: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function getAttributeNodeNS(const namespaceURI, localName: DOMString):
      IDOMAttr; safecall;
    function setAttributeNodeNS(const newAttr: IDOMAttr): IDOMAttr; safecall;
    function getElementsByTagNameNS(const namespaceURI,
      localName: DOMString): IDOMNodeList; safecall;
    function hasAttribute(const name: DOMString): WordBool; safecall;
    function hasAttributeNS(const namespaceURI, localName: DOMString): WordBool; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure normalize; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
  public
    property NativeElement: TdomElement read GetNativeElement;
  end;

{ Tox31DOMText }

  Tox31DOMText = class(Tox31DOMCharacterData, IDOMText)
  protected
    function splitText(offset: Integer): IDOMText; safecall;
  end;

{ Tox31DOMComment }

  Tox31DOMComment = class(Tox31DOMCharacterData, IDOMComment)
  end;

{ Tox31DOMCDATASection }

  Tox31DOMCDATASection = class(Tox31DOMText, IDOMCDATASection)
  end;

{ Tox31DOMDocumentType }

  Tox31DOMDocumentTypeChildren = class;

  Tox31DOMDocumentType = class(Tox31DOMNode, IDOMDocumentType)
  private
    FWrapperDocumentTypeChildren: Tox31DOMDocumentTypeChildren;
    FEntities: IDOMNamedNodeMap;
    FNotations: IDOMNamedNodeMap;
    function GetNativeDocumentType: TdomDocumentType;
  protected
    function get_childNodes: IDOMNodeList; override; safecall;
    function hasChildNodes: WordBool; override; safecall;
    { IDOMDocumentType }
    function get_name: DOMString; safecall;
    function get_entities: IDOMNamedNodeMap; safecall;
    function get_notations: IDOMNamedNodeMap; safecall;
    function get_publicId: DOMString; safecall;
    function get_systemId: DOMString; safecall;
    function get_internalSubset: DOMString; safecall;
  public
    constructor Create(ANativeNode: TdomNode; WrapperDocument: Tox31DOMDocument); override;
    destructor Destroy; override;
    property NativeDocumentType: TdomDocumentType read GetNativeDocumentType;
  end;

  Tox31DOMDocumentTypeChildren = class(TInterfacedObject, IDOMNodeList)
  private
    FWrapperOwnerDocumentType: Tox31DOMDocumentType;
  protected
    { IDOMNodeList }
    function get_item(index: Integer): IDOMNode; safecall;
    function get_length: Integer; safecall;
  public
    constructor Create(NativeDocumentType: Tox31DOMDocumentType);
  end;

{ Tox31DOMNotation }

  Tox31DOMNotation = class(Tox31DOMNode, IDOMNotation)
  private
    function GetNativeNotation: TdomNotation;
  protected
    { IDOMNotation }
    function get_publicId: DOMString; safecall;
    function get_systemId: DOMString; safecall;
  public
    property NativeNotation: TdomNotation read GetNativeNotation;
  end;

{ Tox31DOMEntity }

  Tox31DOMEntity = class(Tox31DOMNode, IDOMEntity)
  private
    function GetNativeEntity: TdomEntity;
  protected
    { IDOMEntity }
    function get_publicId: DOMString; safecall;
    function get_systemId: DOMString; safecall;
    function get_notationName: DOMString; safecall;
  public
    property NativeEntity: TdomEntity read GetNativeEntity;
  end;

{ Tox31DOMEntityReference }

  Tox31DOMEntityReference = class(Tox31DOMNode, IDOMEntityReference)
  end;

{ Tox31DOMProcessingInstruction }

  Tox31DOMProcessingInstruction = class(Tox31DOMNode, IDOMProcessingInstruction)
  private
    function GetNativeProcessingInstruction: TdomProcessingInstruction;
  protected
    { IDOMProcessingInstruction }
    function get_target: DOMString; safecall;
    function get_data: DOMString; safecall;
    procedure set_data(const value: DOMString); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
  public
    property NativeProcessingInstruction: TdomProcessingInstruction
      read GetNativeProcessingInstruction;
  end;

{ Tox31DOMEntity }

  Tox31DOMXPathNamespace = class(Tox31DOMNode, IDOMNode)
  private
    FNativeXPathNamespaceNode: TdomXPathNamespace;
    function GetNativeXPathNamespace: TdomXPathNamespace;
  protected
    // IDomNode
    function get_nodeName: DOMString; override; safecall;
    function get_nodeValue: DOMString; override; safecall;
    procedure set_nodeValue(value: DOMString); override; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_namespaceURI: DOMString; override; safecall;
    function get_prefix: DOMString; override; safecall;
    function get_localName: DOMString; override; safecall;
  public
    constructor Create(ANativeNode: TdomNode; AWrapperDocument: Tox31DOMDocument); override;
    destructor Destroy; override;
    property NativeXPathNamespace: TdomXPathNamespace read GetNativeXPathNamespace;
  end;

{ Tox31DOMDocumentFragment }

  Tox31DOMDocumentFragment = class(Tox31DOMNode, IDOMDocumentFragment)
  end;

{ Tox31DOMDocument }

  Tox31DOMDocument = class(Tox31DOMNode, IDOMDocument, IDOMParseOptions,
    IDOMPersist, IDOMParseError, IDOMXMLProlog)
  private
    FWrapperDOMImpl: Tox31DOMImplementation;
    FDocIsOwned: Boolean;
    FParseError: TParseErrorInfo;
    FPreserveWhitespace: Boolean;
    FDocumentElement: IDOMElement;
    FNativeDocumentElement: TdomElement;
  protected
    function GetNativeDocument: TdomDocument;
    procedure RemoveWhiteSpaceNodes;
    { IDOMDocument }
    function get_doctype: IDOMDocumentType; safecall;
    function get_domImplementation: IDOMImplementation; safecall;
    function get_documentElement: IDOMElement; safecall;
    procedure set_documentElement(const DOMElement: IDOMElement); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function createElement(const tagName: DOMString): IDOMElement; safecall;
    function createDocumentFragment: IDOMDocumentFragment; safecall;
    function createTextNode(const data: DOMString): IDOMText; safecall;
    function createComment(const data: DOMString): IDOMComment; safecall;
    function createCDATASection(const data: DOMString): IDOMCDATASection;
      safecall;
    function createProcessingInstruction(const target,
      data: DOMString): IDOMProcessingInstruction; safecall;
    function createAttribute(const name: DOMString): IDOMAttr; safecall;
    function createEntityReference(const name: DOMString): IDOMEntityReference;
      safecall;
    function getElementsByTagName(const tagName: DOMString): IDOMNodeList;
      safecall;
    function importNode(importedNode: IDOMNode; deep: WordBool): IDOMNode;
      safecall;
    function createElementNS(const namespaceURI,
      qualifiedName: DOMString): IDOMElement; safecall;
    function createAttributeNS(const namespaceURI,
      qualifiedName: DOMString): IDOMAttr; safecall;
    function getElementsByTagNameNS(const namespaceURI,
      localName: DOMString): IDOMNodeList; safecall;
    function getElementById(const elementId: DOMString): IDOMElement; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    { IDOMParseOptions }
    function get_async: Boolean; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_preserveWhiteSpace: Boolean; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_resolveExternals: Boolean; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_validate: Boolean; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_async(Value: Boolean); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_preserveWhiteSpace(Value: Boolean); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_resolveExternals(Value: Boolean); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    procedure set_validate(Value: Boolean); {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    { IDOMPersist }
    function get_xml: DOMString; safecall;
    function asyncLoadState: Integer; safecall;
    function load(source: OleVariant): WordBool; safecall;
    function loadFromStream(const stream: TStream): WordBool; {$IFDEF WRAPVER1.1} overload; {$ENDIF} safecall;
    function loadxml(const Value: DOMString): WordBool; safecall;
    procedure save(destination: OleVariant); safecall;
    procedure saveToStream(const stream: TStream); {$IFDEF WRAPVER1.1} overload; {$ENDIF} safecall;
    procedure set_OnAsyncLoad(const Sender: TObject;
      EventHandler: TAsyncEventHandler); safecall;
{$IFDEF WRAPVER1.1}
    function loadFromStream(const stream: IStream): WordBool; overload; safecall;
    procedure saveToStream(const stream: IStream); overload; safecall;
{$ENDIF}
    { IDOMParseError }
    function get_errorCode: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_url: WideString; safecall;
    function get_reason: WideString; safecall;
    function get_srcText: WideString; safecall;
    function get_line: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_linepos: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    function get_filepos: Integer; {$IFDEF WRAPVER1.1} safecall; {$ENDIF}
    { IDOMXMLProlog }
    function get_Encoding: DOMString; safecall;
    function get_Standalone: DOMString; safecall;
    function get_Version: DOMString; safecall;
    procedure set_Encoding(const Value: DOMString); safecall;
    procedure set_Standalone(const Value: DOMString); safecall;
    procedure set_Version(const Value: DOMString); safecall;
  public
    constructor Create(AWrapperDOMImpl: Tox31DOMImplementation; ANativeDoc: TdomDocument;
      DocIsOwned: Boolean); reintroduce;
    destructor Destroy; override;
    property NativeDocument: TdomDocument read GetNativeDocument;
    property PreserveWhitespace: Boolean read FPreserveWhitespace;
    property WrapperDOMImpl: Tox31DOMImplementation read FWrapperDOMImpl;
  end;

{ Tox31DOMImplementationFactory }

  Tox31DOMImplementationFactory = class(TDOMVendor)
  private
    FGlobalDOMImpl: IDOMImplementation;
  public
    function DOMImplementation: IDOMImplementation; override;
    function Description: String; override;
  end;

var
  OpenXML31Factory: Tox31DOMImplementationFactory;

implementation

uses
  TypInfo,
{$IFDEF LINUX}
  Types,
  IdHttp,
  UrlMon;
{$ENDIF}
{$IFDEF MSWINDOWS}
  Windows,
  ActiveX;
{$ENDIF}

var
  GlobalOx31DOM: Tox31DOMImplementation;

resourcestring
{$IFDEF MSWINDOWS}
  SErrorDownloadingURL = 'Error downloading URL: %s';
  SUrlMonDllMissing = 'Unable to load %s';
{$ENDIF}
  SNodeExpected   = 'NativeNode cannot be null';
  SNotImplemented =
    'This property or method is not implemented in the Open XML Parser';


{ Utility Functions }

function GetNativeNodeOfIntfNode(const Node: IDOMNode): TdomNode;
begin
  if not Assigned(Node) then
    raise DOMException.Create(SNodeExpected);
  Result := (Node as Iox31DOMNodeRef).GetNativeNode;
end;

function MakeNode(NativeNode: TdomNode; WrapperDocument: Tox31DOMDocument): IDOMNode;
const
  NodeClasses: array [xdom_3_1.TDomNodeType] of Tox31DOMNodeClass =
    (Tox31DOMNode, Tox31DOMElement, Tox31DOMAttr, Tox31DOMText, Tox31DOMCDATASection,
     Tox31DOMEntityReference, Tox31DOMEntity, Tox31DOMProcessingInstruction,
     Tox31DOMComment, Tox31DOMDocument, Tox31DOMDocumentType, Tox31DOMDocumentFragment,
     Tox31DOMNotation, Tox31DOMNode, Tox31DOMXPathNamespace);
begin
  if NativeNode <> nil then
    Result := NodeClasses[NativeNode.NodeType].Create(NativeNode, WrapperDocument)
  else
    Result := nil;
end;

function MakeNodeList(const NativeNodeList: TdomNodeList;
  WrapperDocument: Tox31DOMNode): IDOMNodeList; overload;
begin
  Result := Tox31DOMNodeList.Create(NativeNodeList, WrapperDocument);
end;

function MakeNodeList(const NativeXpath: TXpathExpression;
  WrapperDocument: Tox31DOMNode): IDOMNodeList; overload;
begin
  Result := Tox31DOMNodeList.Create(NativeXpath, WrapperDocument);
end;

function MakeNamedNodeMap(const NativeNamedNodeMap: TdomNamedNodeMap; WrapperDocument: Tox31DOMNode):
  IDOMNamedNodeMap;
begin
  Result := Tox31DOMNamedNodeMap.Create(NativeNamedNodeMap, WrapperDocument);
end;

{ Tox31DOMInterface }

function Tox31DOMInterface.SafeCallException(ExceptObject: TObject;
  ExceptAddr: Pointer): HRESULT;
const
  E_FAIL = HRESULT($80004005);
{$IFDEF MSWINDOWS}
var
  HelpFile: string;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  if ExceptObject is EOleException then
    HelpFile := (ExceptObject as EOleException).HelpFile;
  Result := HandleSafeCallException(ExceptObject, ExceptAddr, IDOMNode, '', Helpfile);
{$ELSE}
{$IFDEF LINUX}
  if ExceptObject is Exception then
    SysUtils.SetSafeCallExceptionMsg(Exception(ExceptObject).Message);
  Result := E_FAIL;
{$ELSE}
  DOMVendorNotSupported('SafeCallException', SOpenXML); { Do not localize }
{$ENDIF}
{$ENDIF}
end;

{ TDOMIStreamAdapter }

type
  TDOMIStreamAdapter = class(TStream)
  private
    FStream: IStream;
  public
    constructor Create(const Stream: IStream);
    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
    function Seek(Offset: Longint; Origin: Word): Longint; override;
  end;

constructor TDOMIStreamAdapter.Create(const Stream: IStream);
begin
  inherited Create;
  FStream := Stream;
end;

function TDOMIStreamAdapter.Read(var Buffer; Count: Longint): Longint;
begin
  FStream.Read(@Buffer, Count, @Result)
end;

function TDOMIStreamAdapter.Write(const Buffer; Count: Longint): Longint;
begin
  FStream.Write(@Buffer, Count, @Result);
end;

function TDOMIStreamAdapter.Seek(Offset: Longint; Origin: Word): Longint;
var
  Pos: Largeint;
begin
  FStream.Seek(Offset, Origin, Pos);
  Result := Longint(Pos);
end;

{ Tox31DOMImplementation }

constructor Tox31DOMImplementation.Create;
begin
  inherited;
  FNativeDOMImpl := TDomImplementation.Create(nil);
end;

function Tox31DOMImplementation.createDocument(const namespaceURI,
  qualifiedName: DOMString; doctype: IDOMDocumentType): IDOMDocument;
var
  domDoc: TdomDocument;
  intf: Iox31DOMNodeRef;
  domDocType: TdomDocumentType;
begin
  if Supports(doctype, Iox31DOMNodeRef, intf) then
    domDocType := intf.GetNativeNode as TdomDocumentType
  else
    domDocType := nil;

  if (qualifiedName = '') and (namespaceURI = '') then
  begin
    domDoc := NativeDOMImpl.CreateDoc;
    domDoc.clear;
  end
  else
    domDoc := NativeDOMImpl.CreateDocumentNS(namespaceURI, qualifiedName, domDocType);

  Result := Tox31DOMDocument.Create(self, domDoc, True);
end;

function Tox31DOMImplementation.createDocumentType(const qualifiedName,
  publicId, systemId: DOMString): IDOMDocumentType;
var
  domDocType: TdomDocumentType;
begin
  domDocType := NativeDOMImpl.createDocumentType(qualifiedName, publicId, systemId);
  Result := Tox31DomDocumentType.Create(domDocType, nil);
end;

destructor Tox31DOMImplementation.Destroy;
begin
  inherited;
//  FreeAndNil(FXMLAgent);
end;

procedure Tox31DOMImplementation.FreeDocument(var Doc: TdomDocument);
begin
  FNativeDOMImpl.freeDocument(Doc);
end;

function Tox31DOMImplementation.GetNativeDOMImpl: TdomImplementation;
begin
  Result := FNativeDOMImpl;
end;

function Tox31DOMImplementation.hasFeature(const feature, version: DOMString):
  WordBool;
begin
  Result := NativeDOMImpl.hasFeature(feature, version);
end;

procedure Tox31DOMImplementation.InitParserAgent;
begin
  if not Assigned(FParser) then
  begin
    FParser := TXmlToDomParser.Create(nil);
    FBuilder := TXmlDomBuilder.Create(nil);
    FReader := TXmlStandardDomReader.Create(nil);
    FXpath := TXPathExpression.Create(nil);

    // Setup first step, parsing from file.
    FParser.DOMImpl := NativeDOMImpl;

    // Setup link for second step, parsing to namespace-aware document.
    FBuilder.KeepDocumentTypeDecl := True;
    FBuilder.BuildNamespaceTree := True;
    FReader.NextHandler := FBuilder;
    FReader.PrefixMapping := True;

    // Error handling.
    FReader.OnError := ParseErrorHandler;
  end;
end;

function Tox31DOMImplementation.loadFromStream(const stream: TStream;
  const WrapperDoc: Tox31DOMDocument; var ParseError: TParseErrorInfo): WordBool;
var
  docTemp: TDomDocument;
begin
  FParseError := @ParseError;
  ParseError.errorCode := 0;

  InitParserAgent;

  WrapperDoc.NativeDocument.clear;
  FBuilder.referenceNode := WrapperDoc.NativeDocument;
  docTemp := nil;
  try

    try
      docTemp := FParser.parseStream(stream, '', '', nil) as TDomDocument;
      if ParseError.errorCode = 0 then
        FReader.parse(docTemp);
      Result := (ParseError.errorCode = 0);
    except
      on e: Exception do
      begin
        WrapperDoc.NativeDocument.clear;
        NativeDomImpl.freeUnusedASModelsNS;
        ParseError.reason := e.Message;
        Result := False;
      end;
    end;

  finally
    NativeDomImpl.freeDocument(docTemp);
    NativeDomImpl.freeUnusedASModelsNS;
  end;

  if not WrapperDoc.PreserveWhitespace then
    WrapperDoc.RemoveWhiteSpaceNodes;
end;

function Tox31DOMImplementation.loadxml(const Value: DOMString;
  const WrapperDoc: Tox31DOMDocument; var ParseError: TParseErrorInfo): WordBool;
var
  docTemp: TDomDocument;
begin
  FParseError := @ParseError;
  ParseError.errorCode := 0;

  InitParserAgent;

  WrapperDoc.NativeDocument.clear;
  FBuilder.referenceNode := WrapperDoc.NativeDocument;
  docTemp := nil;
  try

    try
      docTemp := FParser.parseWideString(Value, '', '', nil) as TDomDocument;
      if ParseError.errorCode = 0 then
        FReader.parse(docTemp);
      Result := (ParseError.errorCode = 0);
    except
      on e: Exception do
      begin
        WrapperDoc.NativeDocument.clear;
        NativeDomImpl.freeUnusedASModelsNS;
        ParseError.reason := e.Message;
        Result := False;
      end;
    end;

  finally
    NativeDomImpl.freeDocument(docTemp);
    NativeDomImpl.freeUnusedASModelsNS;
  end;

  if not WrapperDoc.PreserveWhitespace then
    WrapperDoc.RemoveWhiteSpaceNodes;
end;

procedure Tox31DOMImplementation.ParseErrorHandler(sender: TObject; error: TdomError);
begin
  if (error.Severity = DOM_SEVERITY_FATAL_ERROR) or (FParseError.errorCode = 0) then
    with FParseError^ do
    begin
      errorCode := Integer(error.RelatedException);
      errorCodeStr := GetEnumName(TypeInfo(TXmlErrorType), Integer(error.RelatedException));
      // todo: Maybe handle cleartext error messages? reason := e.getErrorStr(iso639_en);
      srcText := error.code;
      line := error.startLineNumber;
      linePos := error.startColumnNumber;
      filePos := error.StartByteNumber;
      url := error.Uri;
    end;
end;

{ Tox31DOMNode }

function Tox31DOMNode.AllocParser: TDomToXmlParser;
begin
  Result := nil;
  if not Assigned(WrapperDocument)
    or not Assigned(WrapperDocument.WrapperDomImpl)
    or not Assigned(WrapperDocument.WrapperDomImpl.NativeDomImpl) then
    Exit;

  Result := TDomToXmlParser.create(nil);
  Result.DOMImpl := WrapperDocument.WrapperDomImpl.NativeDomImpl;

  // Must be freed by calling routine.
end;

constructor Tox31DOMNode.Create(ANativeNode: TdomNode; AWrapperDocument: Tox31DOMDocument);
begin
  FNativeNode := ANativeNode;
  FWrapperDocument := AWrapperDocument;
  inherited Create;
end;

destructor Tox31DOMNode.Destroy;
begin
  inherited;
  FNativeNode := nil;
  FWrapperDocument := nil;
end;

function Tox31DOMNode.appendChild(const newChild: IDOMNode): IDOMNode;
var
  xdnNewChild,
  xdnReturnedChild: TdomNode;
begin
  xdnNewChild := GetNativeNodeOfIntfNode(newChild);
  xdnReturnedChild := NativeNode.appendChild(xdnNewChild);
  if xdnReturnedChild = xdnNewChild then
    Result := newChild else
    Result := MakeNode(xdnReturnedChild, FWrapperDocument);
end;

function Tox31DOMNode.cloneNode(deep: WordBool): IDOMNode;
begin
  Result := MakeNode(NativeNode.CloneNode(deep), FWrapperDocument);
end;

function Tox31DOMNode.get_attributes: IDOMNamedNodeMap;
begin
  if not Assigned(FAttributes) and Assigned(NativeNode.Attributes) then
    FAttributes := MakeNamedNodeMap(NativeNode.Attributes, Self);
  Result := FAttributes;
end;

function Tox31DOMNode.get_childNodes: IDOMNodeList;
begin
  if not Assigned(FChildNodes) then
    FChildNodes := MakeNodeList(NativeNode.ChildNodes, Self);
  Result := FChildNodes;
end;

function Tox31DOMNode.get_firstChild: IDOMNode;
begin
  Result := MakeNode(NativeNode.FirstChild, FWrapperDocument);
end;

function Tox31DOMNode.get_lastChild: IDOMNode;
begin
  Result := MakeNode(NativeNode.LastChild, FWrapperDocument);
end;

function Tox31DOMNode.get_localName: DOMString;
begin
  Result := NativeNode.LocalName;
end;

function Tox31DOMNode.get_namespaceURI: DOMString;
begin
  Result := NativeNode.NamespaceURI;
end;

function Tox31DOMNode.get_nextSibling: IDOMNode;
begin
  Result := MakeNode(NativeNode.NextSibling, FWrapperDocument);
end;

function Tox31DOMNode.get_nodeName: DOMString;
begin
  Result := NativeNode.NodeName;
end;

function Tox31DOMNode.get_nodeType: DOMNodeType;
const
  NodeTypes: array [TDomNodeType] of Word =
   (0, ELEMENT_NODE, ATTRIBUTE_NODE, TEXT_NODE, CDATA_SECTION_NODE,
   ENTITY_REFERENCE_NODE, ENTITY_NODE, PROCESSING_INSTRUCTION_NODE,
   COMMENT_NODE, DOCUMENT_NODE, DOCUMENT_TYPE_NODE, DOCUMENT_FRAGMENT_NODE,
   NOTATION_NODE, 0, ATTRIBUTE_NODE);
begin
  Result := NodeTypes[NativeNode.NodeType];
end;

function Tox31DOMNode.get_nodeValue: DOMString;
begin
  Result := NativeNode.NodeValue;
end;

function Tox31DOMNode.get_ownerDocument: IDOMDocument;
begin
  if not Assigned(FOwnerDocument) then
    FOwnerDocument := Tox31DOMDocument.Create(GlobalOx31DOM, NativeNode.OwnerDocument, False);
  Result := FOwnerDocument;
end;

function Tox31DOMNode.get_parentNode: IDOMNode;
begin
  Result := MakeNode(NativeNode.ParentNode, FWrapperDocument);
end;

function Tox31DOMNode.get_prefix: DOMString;
begin
  Result := NativeNode.Prefix;
end;

function Tox31DOMNode.get_previousSibling: IDOMNode;
begin
  Result := MakeNode(NativeNode.PreviousSibling, FWrapperDocument);
end;

function Tox31DOMNode.GetNativeNode: TdomNode;
begin
  Result := NativeNode;
end;

function Tox31DOMNode.hasChildNodes: WordBool;
begin
  Result := NativeNode.HasChildNodes;
end;

function Tox31DOMNode.insertBefore(const newChild, refChild: IDOMNode): IDOMNode;
begin
  Result := MakeNode(NativeNode.InsertBefore(GetNativeNodeOfIntfNode(newChild),
    GetNativeNodeOfIntfNode(refChild)), FWrapperDocument);
end;

procedure Tox31DOMNode.normalize;
begin
  NativeNode.normalize;
end;

function Tox31DOMNode.removeChild(const childNode: IDOMNode): IDOMNode;
begin
  Result := MakeNode(NativeNode.RemoveChild(GetNativeNodeOfIntfNode(childNode)), FWrapperDocument);
end;

function Tox31DOMNode.replaceChild(const newChild, oldChild: IDOMNode): IDOMNode;
begin
  Result := MakeNode(NativeNode.ReplaceChild(GetNativeNodeOfIntfNode(newChild),
    GetNativeNodeOfIntfNode(oldChild)), FWrapperDocument);
end;

procedure Tox31DOMNode.set_nodeValue(value: DOMString);
begin
  NativeNode.NodeValue := value;
end;

function Tox31DOMNode.supports(const feature, version: DOMString): WordBool;
begin
  Result := NativeNode.supports(feature, version);
end;

{ IDOMNodeSelect }

function Tox31DOMNode.selectNode(const nodePath: WideString): IDOMNode;
var
  xpath: TXpathExpression;
begin
  Result := nil;
  if not Assigned(WrapperDocument) or not Assigned(WrapperDocument.WrapperDOMImpl) then
    Exit;

  xpath := WrapperDocument.WrapperDOMImpl.FXpath;

  xpath.ContextNode := NativeNode;
  xpath.Expression := nodePath;

  if xpath.evaluate and xpath.hasNodeSetResult then
    Result := MakeNode(xpath.resultNode(0), WrapperDocument);
end;

function Tox31DOMNode.selectNodes(const nodePath: WideString): IDOMNodeList;
var
  xpath: TXpathExpression;
begin
  Result := nil;
  if not Assigned(WrapperDocument) or not Assigned(WrapperDocument.WrapperDOMImpl) then
    Exit;

  xpath := WrapperDocument.WrapperDOMImpl.FXpath;

  xpath.ContextNode := NativeNode;
  xpath.Expression := nodePath;

  if xpath.evaluate and xpath.hasNodeSetResult then
    Result := MakeNodeList(xpath, WrapperDocument);
end;

{ IDOMNodeEx Interface }

function Tox31DOMNode.get_text: DOMString;
begin
  Result := NativeNode.TextContent;
end;

procedure Tox31DOMNode.set_text(const Value: DOMString);
var
  Index: Integer;
begin
  for Index := 0 to NativeNode.ChildNodes.Length - 1 do
    NativeNode.RemoveChild(NativeNode.ChildNodes.Item(Index));
  NativeNode.AppendChild(NativeNode.OwnerDocument.CreateTextNode(Value))
end;

function Tox31DOMNode.get_xml: DOMString;
var
  parser: TDomToXmlParser;
begin
  Result := '';
  parser := AllocParser;
  if not Assigned(parser) then
    Exit;
  try

    if not parser.writeToWideString(NativeNode, Result) then
      Result := '';

  finally
    parser.Free;
  end;
end;

procedure Tox31DOMNode.transformNode(const stylesheet: IDOMNode;
  var output: WideString);
begin
  DOMVendorNotSupported('transformNode', sXdom31Xml); { Do not localize }
end;

procedure Tox31DOMNode.transformNode(const stylesheet: IDOMNode;
  const output: IDOMDocument);
begin
  DOMVendorNotSupported('transformNode', sXdom31Xml); { Do not localize }
end;

{ Tox31DOMNodeList }

constructor Tox31DOMNodeList.Create(ANativeNodeList: TdomNodeList;
  AWrapperOwnerNode: Tox31DOMNode);
begin
  inherited Create;
  FNativeNodeList := ANativeNodeList;
  FNativeXpathNodeSet := nil;
  FWrapperOwnerNode := AWrapperOwnerNode;
end;

constructor Tox31DOMNodeList.Create(AnXpath: TXpathExpression;
  AWrapperOwnerNode: Tox31DOMNode);
begin
  FNativeNodeList := nil;
  FNativeXpathNodeSet := AnXpath.acquireXPathResult(TdomXPathNodeSetResult);
  FWrapperOwnerNode := AWrapperOwnerNode;
end;

destructor Tox31DOMNodeList.Destroy;
begin
  FNativeXpathNodeSet.Free;
  inherited;
end;

function Tox31DOMNodeList.get_item(index: Integer): IDOMNode;
begin
  if Assigned(NativeNodeList) then
    Result := MakeNode(NativeNodeList.Item(index), FWrapperOwnerNode.WrapperDocument)
  else
    Result := MakeNode(FNativeXpathNodeSet.item(index), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNodeList.get_length: Integer;
begin
  if Assigned(NativeNodeList) then
    Result := NativeNodeList.Length
  else
    Result := FNativeXpathNodeSet.length;
end;

{ Tox31DOMNamedNodeMap }

constructor Tox31DOMNamedNodeMap.Create(ANativeNamedNodeMap: TdomNamedNodeMap;
  AWrapperOwnerNode: Tox31DOMNode);
begin
  inherited Create;
  FNativeNamedNodeMap := ANativeNamedNodeMap;
  FWrapperOwnerNode := AWrapperOwnerNode;
end;

procedure Tox31DOMNamedNodeMap.CheckNamespaceAware;
begin
  if not NativeNamedNodeMap.namespaceAware then
    raise Exception.Create('NamedNodeMap must be namespace-aware.');
end;

function Tox31DOMNamedNodeMap.get_item(index: Integer): IDOMNode;
begin
  Result := MakeNode(NativeNamedNodeMap.Item(index), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.get_length: Integer;
begin
  Result := NativeNamedNodeMap.Length;
end;

function Tox31DOMNamedNodeMap.getNamedItem(const name: DOMString): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.GetNamedItemNS('', name), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.getNamedItemNS(const namespaceURI,
  localName: DOMString): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.GetNamedItemNS(namespaceURI, localName), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.removeNamedItem(const name: DOMString): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.RemoveNamedItemNS('', name), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.removeNamedItemNS(const namespaceURI,
  localName: DOMString): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.RemoveNamedItemNS(namespaceURI, localName), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.setNamedItem(const arg: IDOMNode): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.SetNamedItemNS(
    GetNativeNodeOfIntfNode(arg)), FWrapperOwnerNode.WrapperDocument);
end;

function Tox31DOMNamedNodeMap.setNamedItemNS(const arg: IDOMNode): IDOMNode;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeNamedNodeMap.SetNamedItemNS(
    GetNativeNodeOfIntfNode(arg)), FWrapperOwnerNode.WrapperDocument);
end;

{ Tox31DOMCharacterData }

function Tox31DOMCharacterData.GetNativeCharacterData: TdomCharacterData;
begin
  Result := NativeNode as TdomCharacterData;
end;

procedure Tox31DOMCharacterData.appendData(const data: DOMString);
begin
  NativeCharacterData.AppendData(data);
end;

procedure Tox31DOMCharacterData.deleteData(offset, count: Integer);
begin
  NativeCharacterData.DeleteData(offset, count);
end;

function Tox31DOMCharacterData.get_data: DOMString;
begin
  Result := NativeCharacterData.Data;
end;

function Tox31DOMCharacterData.get_length: Integer;
begin
  Result := NativeCharacterData.length;
end;

procedure Tox31DOMCharacterData.insertData(offset: Integer;
  const data: DOMString);
begin
  NativeCharacterData.InsertData(offset, data);
end;

procedure Tox31DOMCharacterData.replaceData(offset, count: Integer;
  const data: DOMString);
begin
  NativeCharacterData.ReplaceData(offset, count, data);
end;

procedure Tox31DOMCharacterData.set_data(const data: DOMString);
begin
  NativeCharacterData.Data := data;
end;

function Tox31DOMCharacterData.substringData(offset, count: Integer): DOMString;
begin
  Result := NativeCharacterData.SubstringData(offset, count);
end;

{ Tox31DOMAttr }

function Tox31DOMAttr.GetNativeAttribute: TdomAttr;
begin
  Result := NativeNode as TdomAttr;
end;

function Tox31DOMAttr.get_name: DOMString;
begin
  Result := NativeAttribute.Name;
end;

function Tox31DOMAttr.get_ownerElement: IDOMElement;
begin
  Result := MakeNode(NativeAttribute.OwnerElement, Self.WrapperDocument) as IDOMElement;
end;

function Tox31DOMAttr.get_specified: WordBool;
begin
  Result := NativeAttribute.Specified;
end;

function Tox31DOMAttr.get_value: DOMString;
begin
  Result := NativeAttribute.Value;
end;

procedure Tox31DOMAttr.set_value(const attributeValue: DOMString);
begin
  NativeAttribute.nodeValue := attributeValue;
end;

{ Tox31DOMElement }

procedure Tox31DOMElement.CheckNamespaceAware;
begin
  if not NativeElement.attributes.namespaceAware then
    raise Exception.Create('Element must be namespace-aware.');
end;

function Tox31DOMElement.GetNativeElement: TdomElement;
begin
  Result := NativeNode as TdomElement;
end;

function Tox31DOMElement.get_tagName: DOMString;
begin
  Result := NativeElement.TagName;
end;

function Tox31DOMElement.getAttribute(const name: DOMString): DOMString;
begin
  Result := NativeElement.getAttributeNSLiteralValue('', name);
end;

function Tox31DOMElement.getAttributeNode(const name: DOMString): IDOMAttr;
begin
  Result := MakeNode(NativeElement.GetAttributeNodeNS('', name), Self.WrapperDocument) as IDOMAttr;
end;

function Tox31DOMElement.getAttributeNodeNS(const namespaceURI,
  localName: DOMString): IDOMAttr;
begin
  Result := MakeNode(NativeElement.GetAttributeNodeNS(namespaceURI, localName), Self.WrapperDocument)
    as IDOMAttr;
end;

function Tox31DOMElement.getAttributeNS(const namespaceURI,
  localName: DOMString): DOMString;
begin
  Result := NativeElement.getAttributeNSLiteralValue(namespaceURI, localName);
end;

function Tox31DOMElement.getElementsByTagName(const name: DOMString):
  IDOMNodeList;
begin
  Result := MakeNodeList(NativeElement.GetElementsByTagName(name), Self.WrapperDocument);
end;

function Tox31DOMElement.getElementsByTagNameNS(const namespaceURI,
  localName: DOMString): IDOMNodeList;
begin
  Result := MakeNodeList(NativeElement.GetElementsByTagNameNS(
    namespaceURI, localName), Self);
end;

function Tox31DOMElement.hasAttribute(const name: DOMString): WordBool;
begin
  Result := NativeElement.hasAttributeNS('', name);
end;

function Tox31DOMElement.hasAttributeNS(const namespaceURI,
  localName: DOMString): WordBool;
begin
  Result := NativeElement.hasAttributeNS(namespaceURI, localName);
end;

procedure Tox31DOMElement.removeAttribute(const name: DOMString);
begin
  NativeElement.RemoveAttributeNS('', name);
end;

function Tox31DOMElement.removeAttributeNode(const oldAttr: IDOMAttr): IDOMAttr;
begin
  Result := MakeNode(NativeElement.RemoveAttributeNode(
    GetNativeNodeOfIntfNode(oldAttr) as TdomAttr), Self.WrapperDocument) as IDOMAttr;
end;

procedure Tox31DOMElement.removeAttributeNS(const namespaceURI,
  localName: DOMString);
begin
  NativeElement.RemoveAttributeNS(namespaceURI, localName);
end;

procedure Tox31DOMElement.setAttribute(const name, value: DOMString);
begin
  CheckNamespaceAware;
  NativeElement.setAttributeNS('', name, value);
end;

function Tox31DOMElement.setAttributeNode(const newAttr: IDOMAttr): IDOMAttr;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeElement.SetAttributeNodeNS(
    GetNativeNodeOfIntfNode(newAttr) as TdomAttr), Self.WrapperDocument) as IDOMAttr;
end;

function Tox31DOMElement.setAttributeNodeNS(const newAttr: IDOMAttr): IDOMAttr;
begin
  CheckNamespaceAware;
  Result := MakeNode(NativeElement.SetAttributeNodeNS(
    GetNativeNodeOfIntfNode(newAttr) as TdomAttr), Self.WrapperDocument) as IDOMAttr;
end;

procedure Tox31DOMElement.setAttributeNS(const namespaceURI, qualifiedName,
  value: DOMString);
begin
  CheckNamespaceAware;
  NativeElement.SetAttributeNS(namespaceURI, qualifiedName, value);
end;

procedure Tox31DOMElement.normalize;
begin
  NativeElement.normalize;
end;

{ Tox31DOMText }

function Tox31DOMText.splitText(offset: Integer): IDOMText;
begin
  Result := MakeNode((NativeNode as TdomText).SplitText(offset), Self.WrapperDocument) as IDOMText;
end;

{ Tox31DOMDocumentType }

constructor Tox31DOMDocumentType.Create(ANativeNode: TdomNode; WrapperDocument: Tox31DOMDocument);
begin
  inherited Create(ANativeNode, WrapperDocument);
  FWrapperDocumentTypeChildren := Tox31DOMDocumentTypeChildren.Create(Self);
  FWrapperDocumentTypeChildren._AddRef;
end;

destructor Tox31DOMDocumentType.Destroy;
begin
  FWrapperDocumentTypeChildren._Release;
  inherited Destroy;
end;

function Tox31DOMDocumentType.GetNativeDocumentType: TdomDocumentType;
begin
  Result := NativeNode as TdomDocumentType;
end;

function Tox31DOMDocumentType.get_childNodes: IDOMNodeList;
begin
  Result := FWrapperDocumentTypeChildren;
end;

function Tox31DOMDocumentType.get_entities: IDOMNamedNodeMap;
begin
  if not Assigned(FEntities) then
    FEntities := MakeNamedNodeMap(NativeDocumentType.Entities, Self);
  Result := FEntities;
end;

function Tox31DOMDocumentType.get_internalSubset: DOMString;
begin
  Result := NativeDocumentType.InternalSubset;
end;

function Tox31DOMDocumentType.get_name: DOMString;
begin
  Result := NativeDocumentType.Name;
end;

function Tox31DOMDocumentType.get_notations: IDOMNamedNodeMap;
begin
  if not Assigned(FNotations) then
    FNotations := MakeNamedNodeMap(NativeDocumentType.Notations, Self);
  Result := FNotations;
end;

function Tox31DOMDocumentType.get_publicId: DOMString;
begin
  Result := NativeDocumentType.PublicId;
end;

function Tox31DOMDocumentType.get_systemId: DOMString;
begin
  Result := NativeDocumentType.SystemId;
end;

function Tox31DOMDocumentType.hasChildNodes: WordBool;
begin
  Result := (get_childNodes.Length > 0);
end;

{ Tox31DOMDocumentTypeChildren }

constructor Tox31DOMDocumentTypeChildren.Create(NativeDocumentType: Tox31DOMDocumentType);
begin
  inherited Create;
  FWrapperOwnerDocumentType := NativeDocumentType;
end;

function Tox31DOMDocumentTypeChildren.get_item(index: Integer): IDOMNode;
var
  Len: Integer;
begin
  Len := FWrapperOwnerDocumentType.get_entities.length;
  if index < Len then
    Result := FWrapperOwnerDocumentType.get_entities.item[index]
  else if index < Len + FWrapperOwnerDocumentType.get_notations.length then
    Result := FWrapperOwnerDocumentType.get_notations.item[index - len]
  else
    Result := nil;
end;

function Tox31DOMDocumentTypeChildren.get_length: Integer;
begin
  Result :=
    FWrapperOwnerDocumentType.get_entities.length + FWrapperOwnerDocumentType.get_notations.length;
end;

{ Tox31DOMNotation }

function Tox31DOMNotation.GetNativeNotation: TdomNotation;
begin
  Result := NativeNode as TdomNotation;
end;

function Tox31DOMNotation.get_publicId: DOMString;
begin
  Result := NativeNotation.PublicId;
end;

function Tox31DOMNotation.get_systemId: DOMString;
begin
  Result := NativeNotation.SystemId;
end;

{ Tox31DOMEntity }

function Tox31DOMEntity.GetNativeEntity: TdomEntity;
begin
  Result := NativeNode as TdomEntity;
end;

function Tox31DOMEntity.get_notationName: DOMString;
begin
  Result := NativeEntity.NotationName;
end;

function Tox31DOMEntity.get_publicId: DOMString;
begin
  Result := NativeEntity.PublicId;
end;

function Tox31DOMEntity.get_systemId: DOMString;
begin
  Result := NativeEntity.SystemId;
end;

{ Tox31DOMProcessingInstruction }

function Tox31DOMProcessingInstruction.GetNativeProcessingInstruction:
  TdomProcessingInstruction;
begin
  Result := NativeNode as TdomProcessingInstruction;
end;

function Tox31DOMProcessingInstruction.get_data: DOMString;
begin
  Result := NativeProcessingInstruction.Data;
end;

function Tox31DOMProcessingInstruction.get_target: DOMString;
begin
  Result := NativeProcessingInstruction.Target;
end;

procedure Tox31DOMProcessingInstruction.set_data(const value: DOMString);
begin
  NativeProcessingInstruction.Data := value;
end;

{ Tox31DOMXPathNamespace }

constructor Tox31DOMXPathNamespace.Create(ANativeNode: TdomNode;
  AWrapperDocument: Tox31DOMDocument);
begin
  with ANativeNode as TdomXPathNamespace do
    FNativeXPathNamespaceNode := TdomXPathNamespace.create(nil,
      OwnerElement, namespaceURI, prefix);
  inherited Create(FNativeXPathNamespaceNode, AWrapperDocument);
end;

destructor Tox31DOMXPathNamespace.Destroy;
begin
  inherited;
  FNativeXPathNamespaceNode.Free;
end;

function Tox31DOMXPathNamespace.GetNativeXPathNamespace: TdomXPathNamespace;
begin
  Result := NativeNode as TdomXPathNamespace;
end;

function Tox31DOMXPathNamespace.get_localName: DOMString;
begin
  Result := NativeXPathNamespace.prefix;
end;

function Tox31DOMXPathNamespace.get_namespaceURI: DOMString;
begin
  Result := NativeXPathNamespace.namespaceURI;
end;

function Tox31DOMXPathNamespace.get_nodeName: DOMString;
begin
  Result := 'xmlns:' + NativeXPathNamespace.namespaceURI;
end;

function Tox31DOMXPathNamespace.get_nodeValue: DOMString;
begin
  Result := NativeXPathNamespace.namespaceURI;
end;

function Tox31DOMXPathNamespace.get_prefix: DOMString;
begin
  Result := NativeXPathNamespace.prefix;
end;

procedure Tox31DOMXPathNamespace.set_nodeValue(value: DOMString);
begin
  // Cannot set the value of this Xpath result node.
end;

{ Tox31DOMDocument }

constructor Tox31DOMDocument.Create(AWrapperDOMImpl: Tox31DOMImplementation;
  ANativeDoc: TdomDocument; DocIsOwned: Boolean);
begin
  FDocIsOwned := DocIsOwned;
  FWrapperDOMImpl := AWrapperDOMImpl;
  FPreserveWhitespace := True;
  inherited Create(ANativeDoc, Self);
end;

destructor Tox31DOMDocument.Destroy;
begin
  if FDocIsOwned and Assigned(FWrapperDOMImpl) and Assigned(FNativeNode) then
    FWrapperDOMImpl.FreeDocument(TdomDocument(FNativeNode));
  inherited Destroy;
end;

function Tox31DOMDocument.GetNativeDocument: TdomDocument;
begin
  Result := NativeNode as TdomDocument;
end;

function Tox31DOMDocument.createAttribute(const name: DOMString): IDOMAttr;
begin
  Result := Tox31DOMAttr.Create(NativeDocument.CreateAttributeNS('', name), Self);
end;

function Tox31DOMDocument.createAttributeNS(const namespaceURI,
  qualifiedName: DOMString): IDOMAttr;
begin
  Result := Tox31DOMAttr.Create(NativeDocument.CreateAttributeNS(
    namespaceURI, qualifiedName), Self);
end;

function Tox31DOMDocument.createCDATASection(const data: DOMString):
  IDOMCDATASection;
begin
  Result := Tox31DOMCDATASection.Create(NativeDocument.CreateCDATASection(data), Self);
end;

function Tox31DOMDocument.createComment(const data: DOMString): IDOMComment;
begin
  Result := Tox31DOMComment.Create(NativeDocument.CreateComment(data), Self);
end;

function Tox31DOMDocument.createDocumentFragment: IDOMDocumentFragment;
begin
  Result := Tox31DOMDocumentFragment.Create(NativeDocument.CreateDocumentFragment, Self);
end;

function Tox31DOMDocument.createElement(const tagName: DOMString): IDOMElement;
begin
  Result := Tox31DOMElement.Create(NativeDocument.CreateElement(tagName), Self);
end;

function Tox31DOMDocument.createElementNS(const namespaceURI,
  qualifiedName: DOMString): IDOMElement;
begin
  Result := Tox31DOMElement.Create(
    NativeDocument.CreateElementNS(namespaceURI, qualifiedName), Self);
end;

function Tox31DOMDocument.createEntityReference(const name: DOMString):
  IDOMEntityReference;
begin
  Result := Tox31DOMEntityReference.Create(
    NativeDocument.CreateEntityReference(name), Self);
end;

function Tox31DOMDocument.createProcessingInstruction(const target,
  data: DOMString): IDOMProcessingInstruction;
begin
  Result := Tox31DOMProcessingInstruction.Create(
    NativeDocument.CreateProcessingInstruction(target, data), Self);
end;

function Tox31DOMDocument.createTextNode(const data: DOMString): IDOMText;
begin
  Result := Tox31DOMText.Create(NativeDocument.CreateTextNode(data), Self);
end;

function Tox31DOMDocument.get_doctype: IDOMDocumentType;
begin
  Result := Tox31DOMDocumentType.Create(NativeDocument.docType, Self);
end;

function Tox31DOMDocument.get_documentElement: IDOMElement;
begin
  if not Assigned(FDocumentElement) or
    (FNativeDocumentElement <> NativeDocument.documentElement) then { Test if underlying document NativeElement changed }
  begin
    FNativeDocumentElement := NativeDocument.documentElement;
    FDocumentElement := MakeNode(FNativeDocumentElement, Self) as IDOMElement;
  end;
  Result := FDocumentElement;
end;

function Tox31DOMDocument.get_domImplementation: IDOMImplementation;
begin
  Result := FWrapperDOMImpl;
end;

function Tox31DOMDocument.getElementById(const elementId: DOMString): IDOMElement;
begin
  Result := Tox31DOMElement.Create(NativeDocument.GetElementById(elementId), Self);
end;

function Tox31DOMDocument.getElementsByTagName(const tagName: DOMString): IDOMNodeList;
begin
  Result := MakeNodeList(NativeDocument.GetElementsByTagName(tagName), Self);
end;

function Tox31DOMDocument.getElementsByTagNameNS(const namespaceURI,
  localName: DOMString): IDOMNodeList;
begin
  Result := MakeNodeList(NativeDocument.GetElementsByTagNameNS(
    namespaceURI, localName), Self);
end;

function Tox31DOMDocument.importNode(importedNode: IDOMNode; deep: WordBool):
  IDOMNode;
begin
  Result := MakeNode(NativeDocument.ImportNode(
    GetNativeNodeOfIntfNode(importedNode), deep), Self);
end;

procedure Tox31DOMDocument.set_documentElement(const DOMElement: IDOMElement);
begin
  if Assigned(DOMElement) then
  begin
    if Assigned(NativeDocument.documentElement) then
      NativeDocument.replaceChild(GetNativeNodeOfIntfNode(DOMElement),
        NativeDocument.documentElement)
    else
      NativeDocument.appendChild(GetNativeNodeOfIntfNode(DOMElement));
  end
  else if Assigned(NativeDocument.documentElement) then
      NativeDocument.removeChild(NativeDocument.documentElement);
  FDocumentElement := nil;
end;

{ IDOMParseOptions Interface }

function Tox31DOMDocument.get_async: Boolean;
begin
  Result := False;
end;

function Tox31DOMDocument.get_preserveWhiteSpace: Boolean;
begin
  Result := True;
end;

function Tox31DOMDocument.get_resolveExternals: Boolean;
begin
  Result := False;
end;

function Tox31DOMDocument.get_validate: Boolean;
begin
  Result := False;
end;

procedure Tox31DOMDocument.set_async(Value: Boolean);
begin
  if Value then
    DOMVendorNotSupported('set_async(True)', sXdom31Xml); { Do not localize }
end;

procedure Tox31DOMDocument.set_preserveWhiteSpace(Value: Boolean);
begin
  FPreserveWhitespace := Value;
end;

procedure Tox31DOMDocument.set_resolveExternals(Value: Boolean);
begin
  if Value then
    DOMVendorNotSupported('set_resolveExternals(True)', sXdom31Xml); { Do not localize }
end;

procedure Tox31DOMDocument.set_validate(Value: Boolean);
begin
  if Value then
    DOMVendorNotSupported('set_validate(True)', sXdom31Xml); { Do not localize }
end;

{ IDOMPersist interface }

function Tox31DOMDocument.get_xml: DOMString;
var
  EncodingSave: WideString;
  parser: TDomToXmlParser;
begin
  EncodingSave := NativeDocument.xmlEncoding;
  try
    { We want to be like MSXML and leave the encoding out completely since it's
      going to be unicode because we are returning a unicode string }
    NativeDocument.xmlEncoding := '';

    Result := '';
    parser := AllocParser;
    if not Assigned(parser) then
      Exit;
    try

      if not parser.writeToWideString(NativeDocument, Result) then
        Result := '';

    finally
      parser.Free;
    end;

  finally
    NativeDocument.xmlEncoding := EncodingSave;
  end;
end;

{$IFDEF LINUX}
procedure LoadFromURL(URL: string; Stream: TStream);
var
  IndyHTTP: TIDHttp;
begin
  IndyHTTP := TIDHttp.Create(Nil);
  try
    IndyHttp.Request.Accept := 'text/xml, text/html, application/octet-stream'; { Do not localize }
    IndyHttp.Request.UserAgent := 'Mozilla/3.0 (compatible; Indy Library)'; { Do not localize }
    IndyHttp.Request.ContentType := 'text/xml';   { Do not localize }
    IndyHttp.Request.Location := URL;
    IndyHttp.Request.Connection := URL;
    IndyHttp.Intercept := Nil;
    IndyHTTP.Get(Url, Stream);
  finally
    IndyHTTP.Free;
  end;
end;
{$ENDIF}

{$IFDEF MSWINDOWS}
var
  UrlMonHandle: HMODULE;
  URLDownloadToCacheFile: function(Caller: IUnknown; URL: PAnsiChar;
    FileName: PAnsiChar; FileNameBufLen: DWORD; Reserved: DWORD;
    StatusCB: IInterface {IBindStatusCallback}): HResult; stdcall = nil;

procedure LoadFromURL(URL: string; Stream: TMemoryStream);

  procedure InitURLMon;
  const
    UrlMonLib = 'URLMON.DLL';                             { Do not localize }
    sURLDownloadToCacheFileA = 'URLDownloadToCacheFileA'; { Do not localize }
  begin
    if not Assigned(URLDownloadToCacheFile) then
    begin
      UrlMonHandle := LoadLibrary(UrlMonLib);
      if UrlMonHandle = 0 then
        raise Exception.CreateResFmt(@SUrlMonDllMissing, [UrlMonLib]);
      URLDownloadToCacheFile := GetProcAddress(UrlMonHandle, sURLDownloadToCacheFileA);
    end;
  end;

var
  FileName: array[0..MAX_PATH] of AnsiChar;
begin
  InitURLMon;
  if URLDownloadToCacheFile(nil, PChar(URL), FileName, SizeOf(FileName), 0, nil) <> S_OK then
    raise Exception.CreateResFmt(@SErrorDownloadingURL, [URL]);
  Stream.LoadFromFile(FileName);
end;
{$ENDIF}

function Tox31DOMDocument.load(source: OleVariant): WordBool;
var
  Stream: TMemoryStream;
begin
  if VarType(source) = varOleStr then
  begin
    Stream:= TMemoryStream.create;
    try
      if LowerCase(Copy(Source, 1, 7)) = 'http://' then   { do not localize }
        LoadFromURL(Source, Stream)
      else
        Stream.LoadFromFile(Source);
      result:= loadFromStream(Stream);
    finally
      Stream.free;
    end;
  end
  else
    DOMVendorNotSupported('load(object)', sXdom31Xml); { Do Not Localize }
end;

function Tox31DOMDocument.loadxml(const Value: DOMString): WordBool;
begin
  Result := WrapperDOMImpl.loadxml(Value, self, FParseError);
end;

function Tox31DOMDocument.loadFromStream(const stream: TStream): WordBool;
begin
  Result := WrapperDOMImpl.loadFromStream(stream, Self, FParseError);
end;

procedure Tox31DOMDocument.save(destination: OleVariant);
var
  FStream: TFileStream;
begin
  if VarType(destination) = varOleStr then
  begin
    FStream := TFileStream.Create(destination, fmCreate);
    try
      saveToStream(FStream);
    finally
      FStream.Free;
    end;
  end
  else
    DOMVendorNotSupported('save(object)', sXdom31Xml); { Do Not Localize }
end;

procedure Tox31DOMDocument.saveToStream(const stream: TStream);
var
  parser: TDomToXmlParser;
begin
  parser := AllocParser;
  if not Assigned(parser) then
    Exit;
  try

    if NativeDocument.xmlEncoding = '' then
      parser.writeToStream(NativeDocument, 'UTF-8', stream)
    else
      parser.writeToStream(NativeDocument, NativeDocument.xmlEncoding, stream);

  finally
    parser.Free;
  end;
end;

{$IFDEF WRAPVER1.1}
function Tox31DOMDocument.loadFromStream(const stream: IStream): WordBool;
var
  LStream: TStream;
begin
  LStream := TDOMIStreamAdapter.Create(stream);
  try
    Result := loadFromStream(LStream);
  finally
    LStream.Free;
  end;
end;

procedure Tox31DOMDocument.saveToStream(const stream: IStream);
var
  LStream: TStream;
begin
  LStream := TDOMIStreamAdapter.Create(stream);
  try
    SaveToStream(LStream);
  finally
    LStream.Free;
  end;
end;
{$ENDIF}

{ IDOMParseError }

function Tox31DOMDocument.get_errorCode: Integer;
begin
  Result := FParseError.errorCode;
end;

function Tox31DOMDocument.get_filepos: Integer;
begin
  Result := FParseError.filePos;
end;

function Tox31DOMDocument.get_line: Integer;
begin
  Result := FParseError.line;
end;

function Tox31DOMDocument.get_linepos: Integer;
begin
  Result := FParseError.linePos;
end;

function Tox31DOMDocument.get_reason: WideString;
begin
  Result := FParseError.reason;
end;

function Tox31DOMDocument.get_srcText: WideString;
begin
  Result := FParseError.srcText;
end;

function Tox31DOMDocument.get_url: WideString;
begin
  Result := FParseError.url;
end;

function Tox31DOMDocument.asyncLoadState: Integer;
begin
  result := 0; { Not Supported }
end;

procedure Tox31DOMDocument.set_OnAsyncLoad(const Sender: TObject;
  EventHandler: TAsyncEventHandler);
begin
  DOMVendorNotSupported('set_OnAsyncLoad', sXdom31Xml); { Do Not Localize }
end;

{ IDOMXMLProlog }

function Tox31DOMDocument.get_Encoding: DOMString;
begin
  Result := NativeDocument.xmlEncoding;
end;

procedure Tox31DOMDocument.set_Encoding(const Value: DOMString);
begin
  NativeDocument.xmlEncoding := Value;
end;

function Tox31DOMDocument.get_Standalone: DOMString;
begin
  case NativeDocument.xmlStandalone of
    STANDALONE_YES: Result := 'yes';
    STANDALONE_NO: Result := 'no';
    STANDALONE_UNSPECIFIED: Result := '';
  end;
end;

procedure Tox31DOMDocument.set_Standalone(const Value: DOMString);
begin
  if Value = 'yes' then
    NativeDocument.xmlStandalone := STANDALONE_YES
  else if Value = 'no' then
    NativeDocument.xmlStandalone := STANDALONE_NO
  else
    NativeDocument.xmlStandalone := STANDALONE_UNSPECIFIED;
end;

function Tox31DOMDocument.get_Version: DOMString;
begin
  Result := NativeDocument.xmlVersion;
end;

procedure Tox31DOMDocument.set_Version(const Value: DOMString);
begin
  NativeDocument.xmlVersion := Value;
end;

procedure Tox31DOMDocument.RemoveWhiteSpaceNodes;

  procedure PossiblyDeleteWhiteSpaceNode(dn: TdomNode);
  begin
    if Assigned(dn)
      and (NormalizeWhiteSpace(dn.nodeValue) = '') then
      dn.Free;
  end;

var
  nit: TdomNodeIterator;
  dnPrev, dnCurr: TdomNode;
begin
  get_DocumentElement;
  nit := TdomNodeIterator.create(FNativeDocumentElement, [ntText_Node], nil, True);
  dnPrev := nil;
  try

    repeat
      dnCurr := nit.nextNode;
      if not Assigned(dnCurr) then
        break;

      PossiblyDeleteWhiteSpaceNode(dnPrev);
      dnPrev := dnCurr;
    until False;

    PossiblyDeleteWhiteSpaceNode(dnPrev);

  finally
    nit.Free;
  end;
end;

{ Tox31DOMImplementationFactory }

function Tox31DOMImplementationFactory.DOMImplementation: IDOMImplementation;
begin
  if not Assigned(GlobalOx31DOM) then
  begin
    GlobalOx31DOM := Tox31DOMImplementation.Create;
    FGlobalDOMImpl := GlobalOx31DOM;
  end;
  Result := FGlobalDOMImpl;
end;

function Tox31DOMImplementationFactory.Description: String;
begin
  Result := sXdom31Xml;
end;

initialization
  OpenXML31Factory := Tox31DOMImplementationFactory.Create;
  RegisterDOMVendor(OpenXML31Factory);
finalization
{$IFDEF MSWINDOWS}
  URLDownloadToCacheFile := nil;
  if UrlMonHandle <> 0 then
    FreeLibrary(UrlMonHandle);
{$ENDIF}
  UnRegisterDOMVendor(OpenXML31Factory);
  OpenXML31Factory.Free;
end.
