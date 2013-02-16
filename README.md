FizzBuzzforVBA
==============

FizzBuzzforExcelVBA


社内で、ExcelVBAでFizzBuzzをテーマにTDDのハンズオンのようなものをやってみた時のファイルです。

  * https://github.com/masakitk/FizzBuzzforVBA/

初期状態のファイル：initial_テストVBA.xls
 → サンプルのテストが1つあります
実装後のファイル：テストVBA.xls

となっています。
テスト用の期待結果CSVもあります。
   
## テストの作成
TestClassの指定位置より下に、テストメソッドを追加することでテストを記述します。
検証用のメソッドは3つあります
### AssertEqual
値が同じかどうか
### AssertCellRangeToCSV
セル範囲が、期待結果のCSVファイルと同じかどうか
###AssertCsvValues
2つのCSVファイルの内容が同じかどうか

"" テストの実行
TestRunnerクラスのRunTestsメソッドを実行します。
 コンソールにテスト結果が表示されます。
 test.logファイルにも出力されます。

※普通のExcelでテストツールを使いたい時は、ちょっと面倒ですが5個のクラスモジュールと2個の標準モジュールをインポートすれば使えます

