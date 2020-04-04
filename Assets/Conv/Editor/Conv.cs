using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using UnityEditor;
using UnityEngine;

public class Conv {

	[MenuItem( "Tools/データコンバート" )]
	static void Tools_Excels() {

		var filepath = "Assets/Conv/Book1.xls";

		using( var fs = new FileStream( filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite ) ) {
			IWorkbook book;
			if( Path.GetExtension( filepath ) == ".xls" ) book = new HSSFWorkbook( fs );
			else if( Path.GetExtension( filepath ) == ".xlsx" ) book = new XSSFWorkbook( fs );
			else {
				Debug.LogError( "未対応の拡張子" );
				return;
			}

			for( int sheet_idx = 0, total_sheets = book.NumberOfSheets; sheet_idx < total_sheets; ++sheet_idx ) {
				ISheet sheet = book.GetSheetAt( sheet_idx );

				for( int i = 0, l = sheet.LastRowNum; i <= l; ++i ) {
					var row = sheet.GetRow( i );
					if( row == null ) continue;

					var ss = $"Assets/Conv/Item{i:D2}_{row.Cells[ 0 ].StringCellValue}.asset";
					ItemData item = null;
					if( File.Exists( ss ) ) {
						item = AssetDatabase.LoadAssetAtPath<ItemData>( ss );
					}
					else {
						item = ScriptableObject.CreateInstance<ItemData>();
						AssetDatabase.CreateAsset( item, ss );
					}

					item.NameJP = row.Cells[ 0 ].StringCellValue;
					item.NameEN = row.Cells[ 1 ].StringCellValue;
					item.Price  =  (int)row.Cells[ 2 ].NumericCellValue;
				}
			}

			AssetDatabase.SaveAssets();
			AssetDatabase.Refresh();
		}
	}
}
