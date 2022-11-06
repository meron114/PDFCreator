# PDFCreator
　EXCELとPDFをサマリ画像で簡単に合成できるソフトです。

## 説明
　以下に説明を記載いたします。<br>
### (1)ボタンについて
　 1. 保存先選択・・・合成するPDFの保存先(フォルダ名)を選択できます。<br>
　 2. ファイル一括取得・・・フォルダ内にあるEXCELとPDFファイルのサマリ画像を一括して取得します。<br>
　    ※このとき、EXCEL→PDFの順に取り込まれ最初に取り込んだファイルが置かれているフォルダ名とファイル名が
　　自動で入ります。<br>
　 3. ファイル選択・・・EXCELファイルもしくはPDFファイルを選択し、サマリ画像を表示します。<br>
　 4. 削除・・・選択しているサマリ画像を削除できます。(複数選択可)<br>
　 5. ↶(左回転)・・・選択しているサマリ画像を左に回転します。(複数選択可)<br>
　 6. ↷(右回転)・・・選択しているサマリ画像を右に回転します。(複数選択可)<br>
　 7. クリア・・・サマリ画像とフォルダ名・ファイル名を全てクリアします。<br>
　　  ※ただし、フォルダ名固定もしくはファイル名固定にチェックが入っている場合は、チェックした項目のテキ
　　ストは消されません。<br>
　 8. 保存・・・サマリ画像の順番と向きでPDFを指定された保存先に作成します。<br>
　 9. フォルダ名・・・PDFの保存先のフォルダパスを指定します。<br>
　 10. ファイル名・・・保存するPDFのファイル名を指定します。<br>
　 11. フォルダ名固定・・・クリアした際にフォルダ名のテキストが消されないようにします。<br>
　 12. ファイル名固定・・・クリアした際にファイル名のテキストが消されないようにします。<br>
### (2)その他機能
　・EXCELやPDFをサマリ画像の表示エリアにドラッグアンドドロップすることでもサマリ画像を取り込めます。
　　※取り込んだものはサマリ画像の最後のページから続けて取り込まれます。<br>
　・サマリ画像をドラッグし、位置挿入バーにてドロップすることでページを並び替えることができます。
　　(複数選択可)

## 使い方
　インストーラにてインストールしてご使用下さい。<br>
　[PDFCreatorのインストーラ](installer "installer")<br>
　尚、使用する際はOffice2013のVer以上のEXCELが必要となります。

## Visual studioでの実行について
　Visual studioより実行する場合は以下設定が必要となります。
### (1)参照マネージャーにて以下を追加
　Microsoft Excel 16.0 Object Library<br>
　もしくは、Microsoft Excel 〇〇.〇 Object Library<br>
　※〇〇.〇は使用しているEXCELのバージョンによって変わるため、PCにインストールしているOfficeのバージョン
　　に合わせて設定してください<br>
### (2)Nugetにて以下ライブラリを追加
　・GhostscriptSharp<br>
　・iTextSharp<br>
　・Microsoft-WindowsAPICodePack-Shell<br>
### (3)Resourceにて以下追加
　・[loading_blue.gif](Resources/loading_blue.gif "loading_blue.gif")<br>
　・[loading_green.gif](Resources/loading_green.gif "loading_green.gif")<br>
　・[loading_red.gif](Resources/loading_red.gif "loading_red.gif")<br>
　・[PDF_EXCEL.ico](Resources/PDF_EXCEL.ico "PDF_EXCEL.ico")<br>
### (4)アプリケーションマニフェストファイルを追加し、コメントアウトを削除
　以下の画像を参照してください。<br>
　[アプリケーションマニフェスト_コメントアウトの削除](img/アプリケーションマニフェスト_コメントアウト削除.PNG "アプリケーションマニフェスト_コメントアウト削除")

