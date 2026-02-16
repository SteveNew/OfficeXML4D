unit Office4D.Errors;

interface

uses
  System.SysUtils;

type
  EOfficeDocumentException = class(Exception);

  EPackageException = class(EOfficeDocumentException);
  EPackageNotFound = class(EPackageException);
  EPackageInvalid = class(EPackageException);
  EPartNotFound = class(EPackageException);

  EWordDocumentException = class(EOfficeDocumentException);
  EExcelWorkbookException = class(EOfficeDocumentException);

  ENotImplemented = class(EOfficeDocumentException);

implementation

end.
