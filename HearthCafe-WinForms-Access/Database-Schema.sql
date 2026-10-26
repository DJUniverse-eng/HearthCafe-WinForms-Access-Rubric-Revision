-- Reference schema for Microsoft Access (.accdb).
-- Runtime Db.Initialize() creates these tables automatically where possible.
CREATE TABLE Users (
  Id COUNTER CONSTRAINT PK_Users PRIMARY KEY,
  Username TEXT(80) NOT NULL,
  DisplayName TEXT(120) NOT NULL,
  UserRole TEXT(20) NOT NULL,
  PasswordHash TEXT(128) NOT NULL,
  PasswordSalt TEXT(128) NOT NULL,
  IsActive YESNO NOT NULL
);
CREATE UNIQUE INDEX UX_Users_Username ON Users (Username);

CREATE TABLE MenuItems (
  Id COUNTER CONSTRAINT PK_MenuItems PRIMARY KEY,
  ItemName TEXT(120) NOT NULL,
  Category TEXT(80) NOT NULL,
  Price CURRENCY NOT NULL,
  Description MEMO,
  ImagePath TEXT(255),
  IsAvailable YESNO NOT NULL,
  IsArchived YESNO NOT NULL
);

CREATE TABLE Orders (
  Id COUNTER CONSTRAINT PK_Orders PRIMARY KEY,
  OrderNo TEXT(30) NOT NULL,
  CustomerName TEXT(120),
  OrderNotes MEMO,
  CreatedAt DATETIME NOT NULL,
  CompletedAt DATETIME,
  OrderStatus TEXT(20) NOT NULL,
  Subtotal CURRENCY NOT NULL,
  Discount CURRENCY NOT NULL,
  Total CURRENCY NOT NULL,
  PaymentMethod TEXT(30),
  AmountPaid CURRENCY NOT NULL,
  ChangeAmount CURRENCY NOT NULL,
  CreatedBy TEXT(80),
  IsArchived YESNO NOT NULL
);
CREATE UNIQUE INDEX UX_Orders_OrderNo ON Orders (OrderNo);

CREATE TABLE OrderItems (
  Id COUNTER CONSTRAINT PK_OrderItems PRIMARY KEY,
  OrderId LONG NOT NULL,
  MenuItemId LONG,
  ItemName TEXT(120) NOT NULL,
  UnitPrice CURRENCY NOT NULL,
  Quantity LONG NOT NULL,
  LineTotal CURRENCY NOT NULL
);

CREATE TABLE AppSettings (
  SettingKey TEXT(80) CONSTRAINT PK_AppSettings PRIMARY KEY,
  SettingValue MEMO
);
