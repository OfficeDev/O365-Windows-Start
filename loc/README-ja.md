# Office 365 Starter Project for Windows ストア アプリ #

[日本 (日本語)](https://github.com/OfficeDev/O365-Windows-Start/blob/master/loc/README-ja.md) (日本語)


**目次**

- [概要](#overview)
- [変更履歴](#changehistory)
- [前提条件と構成](#prerequisites)
- [ビルド](#build)
- [プロジェクト関連ファイル](#project)
- [既知の問題](#knownissues)
- [トラブルシューティング](#troubleshooting)
- [その他の技術情報](#additional-resources)
- [質問とコメント](#questions-and-comments)
- [ライセンス](https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-Windows/blob/master/LICENSE.txt)

<a name="overview"></a>

## 概要 ##

Office 365 Starter Project サンプルでは、Office 365 API Tools クライアント ライブラリを使用して、Office 365 の [ファイル]、[カレンダー]、[連絡先] サービスのエンドポイントに対する基本的な操作を説明します。また、サンプルではアプリ 1 つで複数の Office 365 サービスに対する認証を行う方法や、[ユーザーとグループ] サービスからユーザー情報を取得する方法についてご確認いただけます。 このプロジェクトのアップデート時に電子メールなどさらに多くの API の使用方法についての例を追加する予定ですので、引き続きご確認ください。

このサンプルで実行できる操作を次に示します。

予定表  
  - カレンダー イベントを取得する  
  - イベントを作成する  
  - イベントを更新する  
  - イベントを削除する  

連絡先  
  - 連絡先を取得する  
  - 連絡先を作成する  
  - 連絡先を更新する  
  - 連絡先を削除する  
  - 連絡先の写真を変更する  

マイ ファイル  
  - ファイルとフォルダーを取得する  
  - テキスト ファイルを作成する  
  - ファイルまたはフォルダーを削除する  
  - テキスト ファイルのコンテンツを読みこむ (OneDrive)  
  - テキスト ファイルのコンテンツを更新する  
  - ファイルをダウンロードする  
  - ファイルをアップロードする  

ユーザーとグループ  
  - 表示名を取得する  
  - ジョブ タイトルを取得する  
  - プロファイル画像を取得する  
  - ユーザー ID を取得する  
  - サインイン / サインアウト状態を確認する  
	
メール  
  - ページ結果でメールを受信する  
  - メールを送信する  
  - メールを削除する  

<a name="changehistory"></a>
## 変更履歴 ##
2015 年 1 月 26 日:

- メール機能を追加しました


2014 年 12 月 17 日:

- 認証フローと AuthenticationHelper.cs ファイルでのクライアント作成を簡略化

- アプリでサービスが何回も照会されないように検出サービスから情報をキャッシュする機能を追加

- 企業イントラネットと企業アカウントのサポートを追加 `AuthenticationContext` オブジェクトの `UseCorporateNetwork` プロパティは True に設定されています。また、プロジェクトにエンタープライズ認証、プライベート ネットワーク、共有ユーザー証明書機能の宣言を追加しています。詳細については、[アプリ機能の宣言 (Windows ランタイム アプリ)](http://aka.ms/vaha2s) をご覧ください。

<a name="prerequisites"></a>

## 前提条件と構成 ##

このサンプルを実行するには次のものが必要です。  

  - Windows 8.1  
  - Visual Studio 2013 更新プログラム 4。  
  - [Office 365 API Tools バージョン 1.4.50428.2](http://aka.ms/k0534n)。  
  - [Office 365 開発者向けサイト](http://aka.ms/ro9c62)  
  - このプロジェクトの [ファイル] の部分を使用するには、Web ブラウザーから OneDrive for Business に初めてするサインオンする際に、自身のテナントで OneDrive for Business をセットアップする必要があります。

**注:**このサインプルは、Visual Studio 2015 Update 1 を使用している場合に動作します。ただし、サンプルを構成した後は、App.xaml ファイルの中のキーと値のペアを編集する必要があります。 

`<x:String x:Key="ida:ClientId">your client id</x:String>`.

サンプルのコードは `ClientID` キーを探すため、キーと値のペアは次のようになります:

`<x:String x:Key="ida:ClientID">your client id</x:String>`.

###サンプルを構成する

次の手順に従ってサンプルを構成します。

   1. Visual Studio 2013 を使用して O365-APIs-Start-Windows.sln ファイルを開きます。
   2. ソリューションをビルドします。 NuGet パッケージ復元機能で packages.config ファイルにリストされたアセンブリが読み込まれます。 一部のアセンブリが旧バージョンにならないように、次の手順で接続済みサービスを追加する前にこれを行う必要があります。
   3. Office 365 サービスを使用するようアプリを登録し構成します (詳細は以下をご覧ください)。


###Office 365 API を使用するようアプリを登録します。

登録は Office 365 API Tools for Visual Studio で自動的に行うことができます。 必ず Visual Studio ギャラリーから Office 365 API ツールをダウンロードしてインストールしてください。

   1. [ソリューション エクスプローラー] ウィンドウで、[Office365Starter] プロジェクト、[追加]、[接続済みサービス] と選択します。
   2. [サービス マネージャー] ダイアログ ボックスが表示されます。 Office 365 を選択してアプリを登録します。
   3. [サインイン] ダイアログ ボックスで、Office 365 テナント用のユーザー名とパスワードを入力します。 自分の Office 365 開発者向けサイトを使用することをお勧めします。 多くの場合、このユーザー名は <your-name>@<tenant-name>.onmicrosoft.com というパターンになります。 自分の開発者向けサイトを持っていない場合、MSDN 特典の一部として無料で、または無料試用版にサインアップすることで入手できます。 ユーザーはテナント管理ユーザーである必要があることにご注意ください。Office 365 開発者向けサイトの一部として作成されたテナントでは、ほとんどの場合テナント管理ユーザーになります。 また、開発者アカウントは通常 1 つのサインインに制限されます。
   4. サインイン後、すべてのサービスを確認できます。 最初はサービスを使用するようアプリが登録されていないため、権限は選択されません。 
   5. このサンプルで使用するサービスを登録するには、次の権限を選択します。  
   	- (予定表) – 予定表の読み取りと書き込み  
	- (連絡先) － 連絡先の読み取りと書き込み 
	- (マイ ファイル) － ファイルの読み取りと書き込み
	- (ユーザーとグループ) － サインインとプロファイルの読み取り  
	- (Users and Groups) – 所属する組織のディレクトリへのアクセス権限
	- (メール) - メールの読み取りと書き込み
	- (メール) - ユーザーとしてのメールの送信
  
   6. [サービス マネージャー] ダイアログ ボックスで [OK] をクリックします。

**メモ:** 手順 6 でパッケージのインストール中にエラーが発生した場合 (たとえば 「"Microsoft.Azure.ActiveDirectory.GraphClient" が見つかりません」など)、ソリューションを保存したローカル パスが長すぎない / 深すぎないことをご確認ください。 デバイスのルート近くにソリューションを移動すると問題が解決します。 また、今後のアップデートでフォルダー名を短くする予定です。      

<a name="build"></a>
## ビルド ##

Visual Studio にソリューションを読み込ませたら、F5 を押してビルドとデバッグを行います。 ソリューションを実行し、Office 365 組織アカウントでサインインします。

<a name="project"></a>
## プロジェクト関連ファイル ##

**ヘルパー クラス**  
   - CalendarOperations.cs  
   - FileOperations.cs  
   - UserOperations.cs  
   - AuthenticationHelper.cs  
   - ContactOperations.cs  

**ビュー モデル**  
   - CalendarViewModel.cs  
   - EventViewModel.cs  
   - UserViewModel.cs  
   - ContactsViewModel.cs  
   - FilesViewModel.cs  
   - FileSystemItemViewModel.cs  
   - ContactItemViewModel.cs  

<a name="knownissues"></a>
## 既知の問題 ##



- この時点では存在しませんが、何かありましたらお知らせください。 情報をお待ちしています。 

<a name="troubleshooting"></a>
## トラブルシューティング ##



- システム特権を持たずグローバル管理者でない一般ユーザーとして Office 365 に接続すると、「十分な権限がありません」例外が発生します。 接続済みサービスを追加する際に*組織のディレクトリにアクセスする*権限が設定されていることをご確認ください。




- パッケージのインストール中に「Microsoft.Azure.ActiveDirectory.GraphClient" が見つかりません」などのエラーが発生します。 ソリューションを保存したローカル パスが長すぎない / 深すぎないことをご確認ください。 デバイスのルート近くにソリューションを移動すると問題が解決します。 また、今後のアップデートでフォルダー名を短くする予定です。  



- アプリが [[Windows プライバシー設定]](http://aka.ms/gqqx6p) メニュー内のアカウント情報にアクセスできない場合、展開や実行した後に認証エラーが発生します。 **[アプリで自分の名前、画像、その他のドメイン アカウント情報にアクセスすることを許可する]** を **[オン]** に設定します。 この設定は、Windows Update でリセットできます。 

- インストール済みアプリに対して Windows アプリ認定キットを実行すると、そのアプリがサポート済み API のテストに失敗します。 Visual Studio ツールでインストールされた一部のアセンブリが旧バージョンであることが原因の可能性があります。 プロジェクトの packages.config ファイルで Microsoft.Azure.ActiveDirectory.GraphClient と Microsoft.OData アセンブリのエントリをご確認ください。 これらのアセンブリのバージョン番号が、[このリポジトリの packages.config バージョン](https://github.com/OfficeDev/O365-Windows-Start/blob/master/Office365StarterProject/packages.config)のバージョン番号と一致していることをご確認ください。 更新したアセンブリでソリューションをリビルドして再インストールするときは、アプリがサポート済み API テストに合格する必要があります。

## その他の技術情報 ##
* [Office 365 API プラットフォームの概要](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [ファイル REST API リファレンス](https://msdn.microsoft.com/office/office365/api/files-rest-operations)
* [Outlook 予定表 REST API リファレンス](http://msdn.microsoft.com/office/office365/api/calendar-rest-operations)
* [Outlook メール REST API リファレンス](https://msdn.microsoft.com/office/office365/api/mail-rest-operations)
* [Microsoft Office 365 API ツール](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office デベロッパー センター](http://dev.office.com/)
* [Office 365 API スタート プロジェクトおよびサンプル コード](http://msdn.microsoft.com/office/office365/howto/starter-projects-and-code-samples)
* [Windows ストア、電話、およびユニバーサル アプリで Office 365 に接続する](https://github.com/OfficeDev/O365-Win-Connect)
* [Windows 用 Office 365 コード スニペット](https://github.com/OfficeDev/O365-Win-Snippets)
* [Windows 用 Office 365 プロファイル サンプル](https://github.com/OfficeDev/O365-Win-Profile)
* [サイト用 Office 365 REST API Explorer](https://github.com/OfficeDev/Office-365-REST-API-Explorer)


## 質問とコメント

O365 Windows Starter プロジェクトについて、Microsoft にフィードバックをお寄せください。質問や提案につきましては、このリポジトリの「[問題](https://github.com/OfficeDev/O365-Windows-Start/issues)」セクションに送信できます。

Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問またはコメントには、必ず [Office365] および [API] のタグを付けてください。


## 著作権 ##

Copyright (c) Microsoft. All rights reserved.



