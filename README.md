# WorksheetFunction-Regex
Excelの正規表現関数をVBAでWorksheetFunctionとして使う (Using Excel Regular Expression Functions as WorksheetFunctions in VBA)  
初回投稿日：2025年1月17日  
最終更新日：2025年1月22日  

## 1. 概要
Microsoft Excelの正規表現関数（REGEXTEST、REGEXREPLACE、REGEXEXTRACT）は、2024年5月にリリースされました。これらの関数は、Excelの365 Insiderプレビュー版で最初に導入されましたが、現在は一般公開されています。これらの正規表現関数は、Microsoft 365のサブスクリプションを持っているユーザーが利用できます。具体的には、Excel for Microsoft 365、Excel for the web、Excel for iOS、Excel for Androidなどのバージョンで利用可能です。   
これらの関数は、比較的新しいPCRE2（Perl Compatible Regular Expressions）に準拠した、より強力な正規表現の構文を記述できます。しかしこれらの正規表現関数はVBAのWorksheetFunctionオブジェクトのメソッドに出てきません。   
この問題を解決するため、ここではEvaluateメソッドを利用してワークシート関数を呼び出す方法を利用しています。また、パラメータの引き渡しに独自のユーザー定義関数（Arg）を用いることによって、文字列だけでなく配列を含むさまざまなデーター型を引き渡すことができ、汎用性を高めています。

### ここで紹介するワークシート関数
|関数名|  概要  |
|  :---  |  :---  |
|WSF_REGEXTEST|指定されたテキストが正規表現パターンに一致するかどうかを判定します。|   
|WSF_REGEXREPLACE|指定された正規表現パターンに一致する文字列を別の文字列で置換します。|    
|WSF_REGEXEXTRACT|指定された正規表現パターンに一致する文字列を抽出します。|  
    
## 2. 解説   
以下にそれぞれの関数について解説します。   
ソースコードはこのリポジトリーにある [RE_Module_WSF.bas](RE_Module_WSF.bas) をダウンロードしてお使いください。  
      
### 2.1 WSF_REGEXTEST 
テキストの任意の部分が正規表現パターンと一致するかどうか検査します。  
```  
構文： WSF_REGEXTEST(text, pattern, [case_sensitivity])  
```  
各パラメータはワークシート関数（REGEXTEST）に従います。  
textには文字列または配列を指定できます。配列を指定した場合は、戻り値としてそれに対応する配列が返されます。

### 2.2 WSF_REGEXREPLACE
指定されたテキスト内の文字列を、パターンに一致する文字列を置換に置き換えます。  
```  
構文： WSF_REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])  
```
各パラメータはワークシート関数（REGEXREPLACE）に従います。  
textには文字列または配列を指定できます。配列を指定した場合は、戻り値としてそれに対応する配列が返されます。

### 2.3 WSF_REGEXEXTRACT
指定されたテキスト内で正規表現パターンに一致する文字列を抽出します。  
```
構文： WSF_REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])  
```
各パラメータはワークシート関数（REGEXEXTRACT）に従います。  
textには文字列または配列を指定できます。ただし配列を指定した場合は配列の第一要素のみ評価されます。これはREGEXEXTRACT自体の動作によるものです。  

# 3. ライセンス
このコードはMITライセンスに基づき利用できます。  
   
■ 
