<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/942007512/24.2.1%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1280328)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
# Word Processing Document API - Iterate through all Sub-documents in a Document

This example implements an extension method for RichEditDocumentServer.Document and specify a delegate that will be executed for each SubDocument in a document. This approach allows you to modify the entire document content in a single place.

# Implementation Details

RichEditDocumentServer's document is split into logical parts - [SubDocuments](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.SubDocument). The document body, text boxes, [comments](https://docs.devexpress.com/OfficeFileAPI/116863/word-processing-document-api/word-processing-document/comments), [headers and footers](https://docs.devexpress.com/OfficeFileAPI/15310/word-processing-document-api/word-processing-document/headers-and-footers) for different document sections are stored in separate SubDocuments. Each SubDocument contains its own textual content, fields, bookmarks, hyperlinks, images, and shapes. In scenarios when such document elements should be edited for the entire document, it's necessary to get all document SubDocuments and edit each SubDocument separately.

# Files to Review

| C# | Visual Basic |
|---|---|
| [Program.cs](./CS/Program.cs) | [Program.vb](./VB/Program.vb) |
| [SubDocumentHelper.cs](./CS/SubDocumentHelper.cs) | [SubDocumentHelper.vb](./VB/SubDocumentHelper.vb) |

# Documentation

* [Word Processing - Document Model](https://docs.devexpress.com/OfficeFileAPI/15305/word-processing-document-api/word-processing-document/document-structure/document-model)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-iterate-through-all-sub-documents&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-iterate-through-all-sub-documents&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
