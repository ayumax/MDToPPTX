# MDToPPTX
A library that reads markdown format files and outputs PowerPoint format (pptx) files.
It is .Net Standard 2.0 compatible.

[Markdig](https://github.com/lunet-io/markdig) is used for analysis of markdown

# Corresponding syntax
+ HEADERS(Corresponds to Level 1 or Level 2)
+ BLOCKQUOTES
+ LISTS
+ CODE BLOCKS
+ LINKS
+ EMPHASIS
+ Images(1 pixel is arranged as 1 mm)
+ Table

## MarkPP.exe
Executable file for using MDToPPTX on Windows

## How to use MarkPP.exe

Execute the following command

```
MarkPP.exe "markdownfle path" "title" "subtitle" "setting json path"
```
+ markdownfle path:Markdown file path(*.md)
+ title:Title to be described on the first page of pptx
+ subtitle:Sub title to be described on the first page of pptx
+ setting json path:Configuration file path(*.json)

## input Markdown例

Write \ --- on the boundary of the sheet.

```
# テストシート1

テストシートです。  
**太字**です  
*イタリック*も対応  
~~打消しも可能~~

[ハイパーリンク例](http://ayumax.hatenablog.com/)

↓コードブロック

```　　　

class ClassA  
{  
    public ClassA()  
    { 

    }

    public void Func()
    {

    }
}  

```　　　


---

# テストシート2

テストシート2枚目です

箇条書きにも対応

+ 箇条書きです
+ インラインコードも`対応`してます

イメージの挿入もできます
![image1](ayumax.jpg)

引用文はこちら
> 引用サンプル
> です

---
# テストシート3
## サブタイトル

テストシート3枚目です

1. 数値の箇条書き
1. 数値の箇条書き2行目
1. 数値の箇条書き3行目

表も書けます

| Left align | Right align | Center align |
|:-----------|------------:|:------------:|
| 1行目左    | 1行目        | 1行目右      |
| 2行目左    | 2行目        | 2行目右      |
| 3行目左    | 3行目        | 3行目右      |    

```

## output PPTX
![sheet1](images/sheet1.JPG)
![sheet2](images/sheet2.JPG)
![sheet3](images/sheet3.JPG)
