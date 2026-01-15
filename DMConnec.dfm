object dtmConnec: TdtmConnec
  OldCreateOrder = True
  OnCreate = DataModuleCreate
  Height = 426
  Width = 944
  object FDConnection: TFDConnection
    Params.Strings = (
      'User_Name=sysdba'
      'Database=C:\MM\Database\DATABASE - DIGNA.GDB'
      'Password=masterkey'
      'Protocol=TCPIP'
      'Server=localhost'
      'SQLDialect=1'
      'CharacterSet=WIN1252'
      'DriverID=FB')
    LoginPrompt = False
    Transaction = FDTransaction
    UpdateTransaction = FDTransaction
    Left = 40
    Top = 24
  end
  object FDTransaction: TFDTransaction
    Connection = FDConnRemoto
    Left = 32
    Top = 104
  end
  object FDGUIxWaitCursor: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 133
    Top = 110
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    VendorLib = 'fbclient.dll'
    Left = 280
    Top = 96
  end
  object FDQuery: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      'select * from tabconf')
    Left = 24
    Top = 160
  end
  object FDConnRemoto: TFDConnection
    Params.Strings = (
      'User_Name=MMSOFTWARES'
      'Database=/firebird/mmsoftwares.gdb'
      'Password=110550'
      'Protocol=TCPIP'
      'Server=firebird.mmsoftwares.com.br'
      'SQLDialect=1'
      'CharacterSet=WIN1252'
      'DriverID=FB')
    LoginPrompt = False
    Transaction = FDTransaction
    UpdateTransaction = FDTransaction
    Left = 496
    Top = 104
  end
  object FDQryRemoto: TFDQuery
    Connection = FDConnRemoto
    SQL.Strings = (
      'select * from tabconf')
    Left = 496
    Top = 264
  end
  object FDTransRemoto: TFDTransaction
    Connection = FDConnRemoto
    Left = 496
    Top = 184
  end
end
