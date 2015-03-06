{-----------------------------------------------------------------------------
 Unit Name: xmldocxpath
 Author:    Tor Helland
 Purpose:   Use Xpath queries in a simple way on IXmlNode.
 History:   20090809 th Isolated selectNode/selectNodes into this file.

 This code is released under the Mozilla Public License Version 1.1.

 LICENSE:
 The contents of this file are subject to the Mozilla Public License
 Version 1.1 (the "License"); you may not use this file except in
 compliance with the License. You may obtain a copy of the License at
 http://www.mozilla.org/MPL/

 Software distributed under the License is distributed on an "AS IS"
 basis, WITHOUT WARRANTY OF ANY KIND, either express or implied. See the
 License for the specific language governing rights and limitations
 under the License.

 The Original Code is xmldocxpath.pas.

 The Initial Developer of the Original Code is Tor Helland.
 Portions created by Tor Helland are Copyright (C) 2005-2009 Tor Helland.
 All Rights Reserved.

 Alternatively, the contents of this file may be used under the terms
 of the GNU General Public License Version 2 or later (the "GPL"), in which case the
 provisions of GPL are applicable instead of those
 above. If you wish to allow use of your version of this file only
 under the terms of the GPL and not to allow others to use
 your version of this file under the MPL, indicate your decision by
 deleting the provisions above and replace them with the notice and
 other provisions required by the GPL. If you do not delete
 the provisions above, a recipient may use your version of this file
 under either the MPL or the GPL.
-----------------------------------------------------------------------------}
unit xmldocxpath;

interface
uses
  Windows,
  Classes,
  SysUtils,
  xmldom,
  XmlIntf;

// Calling IDomNodeSelect.
function selectNode(xnRoot: IXmlNode; const nodePath: WideString): IXmlNode;
function selectNodes(xnRoot: IXmlNode; const nodePath: WideString): IXMLNodeList;

implementation
uses
  StrUtils,
  XmlDoc;

function selectNode(xnRoot: IXmlNode; const nodePath: WideString): IXmlNode;
var
  intfSelect: IDomNodeSelect;
  dnResult: IDomNode;
  intfDocAccess: IXmlDocumentAccess;
  doc: TXmlDocument;
begin
  Result := nil;
  if not Assigned(xnRoot)
    or not Supports(xnRoot.DOMNode, IDomNodeSelect, intfSelect) then
    Exit;

  dnResult := intfSelect.selectNode(nodePath);
  if Assigned(dnResult) then
  begin
    if Supports(xnRoot.OwnerDocument, IXmlDocumentAccess, intfDocAccess) then
      doc := intfDocAccess.DocumentObject
    else
      doc := nil;
    Result := TXmlNode.Create(dnResult, nil, doc);
  end;
end;

function selectNodes(xnRoot: IXmlNode; const nodePath: WideString): IXMLNodeList;
resourcestring
  rcErrSelectNodesNoRoot = 'Called selectNodes without specifying proper root.';
var
  intfSelect: IDomNodeSelect;
  intfAccess: IXmlNodeAccess;
  dnlResult: IDomNodeList;
  intfDocAccess: IXmlDocumentAccess;
  doc: TXmlDocument;
  i: Integer;
  dn: IDomNode;
begin
  // Always return a node list, even if empty.
  if not Assigned(xnRoot)
    or not Supports(xnRoot, IXmlNodeAccess, intfAccess)
    or not Supports(xnRoot.DOMNode, IDomNodeSelect, intfSelect) then
    XmlDocError(rcErrSelectNodesNoRoot);

  // TXMLNodeList must have an owner (xnRoot).
  Result := TXmlNodeList.Create(intfAccess.GetNodeObject, '', nil);

  dnlResult := intfSelect.selectNodes(nodePath);
  if Assigned(dnlResult) then
  begin
    if Supports(xnRoot.OwnerDocument, IXmlDocumentAccess, intfDocAccess) then
      doc := intfDocAccess.DocumentObject
    else
      doc := nil;

    for i := 0 to dnlResult.length - 1 do
    begin
      dn := dnlResult.item[i];
      Result.Add(TXmlNode.Create(dn, nil, doc));
    end;
  end;
end;

end.

