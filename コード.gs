/**
 * 検索
 */
function search( )
{
  const searchSheet = SpreadsheetApp.openById("1tRSLoWaexEflRHelkITugLJGnUxgIdIg79vdwTfXRdc").getSheetByName( "検索シート");
  const formSheet   = SpreadsheetApp.openById("1tRSLoWaexEflRHelkITugLJGnUxgIdIg79vdwTfXRdc").getSheetByName( "フォームの回答 1");

  //検索データ取得
  let searchColumn = 2, search = [];
  for( ; searchColumn < 11; searchColumn++)
  {
    search.push( searchSheet.getRange( 3, searchColumn).getDisplayValue( ));
  }
  //console.log( search);
  searchSheet.getRange( "B8:J58").clearContent( );

  //フォームデータ配列化
  formSheet.getRange( 'A2').activate();
  const formData = formSheet.getRange( formSheet.getSelection().getNextDataRange( SpreadsheetApp.Direction.DOWN).getA1Notation( ).replace( /A2:A/, 'B2:J')).getDisplayValues( );
  //console.log( formData.length);

  //部分一致検索
  let resultRow = 8;
  for( let formRow = 0; formRow < formData.length; formRow++)
  {
    let hits = [];
//console.log( formSheet.getRange( formRow, 1).getValue( ));
    for( let formColumn = 0; formColumn < formData[ formRow].length; formColumn++)
    {
      if( search[ formColumn] === '')
      {
        continue;
      }
//console.log( search[ formColumn] +", "+ formData[ formRow][ formColumn]);
      if( formData[ formRow][ formColumn].indexOf( search[ formColumn]) >= 0)
      {
        hits.push( formColumn);
      }
    }
    if( hits.length === 0)
    {
      continue;
    }
    //const l = search[ formColumn].lenght;
    for( let resultColumn = 2; resultColumn < 11; resultColumn++)
    {
      if( resultColumn - 2 === hits[ 0])
      {
//console.log( resultColumn - 2 +", "+ hits[ 0]);
//console.log( formData[ formRow][ resultColumn - 2]);
//console.log( search[ hits[ 0]].length);
        //ハイライト
        const rtv = SpreadsheetApp.newRichTextValue( ).setText( formData[ formRow][ resultColumn - 2]);
        let style = [];
        for( let i = 0; formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i) > -1; )
        {
          style.push( [
            formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i)
            , formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i) + search[ hits[ 0]].length
            , SpreadsheetApp.newTextStyle( ).setBold( true).build( )
          ]);
          style.push( [
            formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i)
            , formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i) + search[ hits[ 0]].length
            , SpreadsheetApp.newTextStyle( ).setForegroundColor( 'red').build( )
          ]);
          i = formData[ formRow][ resultColumn - 2].indexOf( search[ hits[ 0]], i) + search[ hits[ 0]].length;
        }
//console.log( style);
        for( let i = 0; i < style.length; i++)
        {
          rtv.setTextStyle( style[ i][ 0], style[ i][ 1], style[ i][ 2]);
        }
//console.log( rtv);
        searchSheet.getRange( resultRow, resultColumn).setRichTextValue( rtv.build( ));
        hits.shift( );
      }
      else
      {
        searchSheet.getRange( resultRow, resultColumn).setValue( formData[ formRow][ resultColumn - 2]);
      }
    }
    resultRow++;
    if( resultRow > 57)
    {
      Browser.msgBox( '５０件を超えました。条件を変更して結果を絞って下さい。');
      return;
    }
  }
}

/**
 * フォーム送信時
 */
function submit( )
{
  //曲名　読み　が空の場合に原題を日本語訳する
  const formSheet = SpreadsheetApp.openById( "1tRSLoWaexEflRHelkITugLJGnUxgIdIg79vdwTfXRdc").getSheetByName( "フォームの回答 1");
  const lastRow = formSheet.getLastRow( );
  const yomiCell = formSheet.getRange( lastRow, 10);
  if( yomiCell.getDisplayValue( ) !== "")
  {
    return;
  }
  yomiCell.setValue( LanguageApp.translate( formSheet.getRange( lastRow, 9).getDisplayValue( ), "", "ja"));
}

