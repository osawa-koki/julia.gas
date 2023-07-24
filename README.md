# julia.gas

🕵️‍♀️🕵️‍♀️🕵️‍♀️ GASを使用してGoogleスプレッドシート上にジュリア集合を描画してみる！  

![成果物](./docs/img/fruit.gif)  

## 準備方法

### 1. Google Apps Script APIの有効化

[Google Apps Script API](https://script.google.com/home/usersettings)を有効化します。  

### 2. Claspのインストール

以下のコマンドでClaspをインストールします。  

```shell
yarn global add clasp
```

### 3. Claspのログイン

以下のコマンドでログインします。  
ブラウザが立ち上がり、Googleアカウントの認証を行います。  

```shell
clasp login
```

### 4. プロジェクトの作成

Node.jsプロジェクトを初期化します。  
また、GASの型定義ファイルをインストールします。  

```shell
yarn init
yarn add -D @types/google-apps-script
```

### 5. Claspプロジェクトの作成

以下のコマンドでClaspプロジェクトを作成します。  

```shell
clasp create
```

### 6. その他イロイロのコマンド

```shell
# ローカルのファイルをGASにアップロード
clasp push

# GASのファイルをローカルにダウンロード
clasp pull

# GASのファイルを開く
clasp open
```

### 7. プロパティサービスへのデータの登録

`https://script.google.com/u/0/home/projects/⭐️プロジェクトID⭐️/settings`から登録可能です。  
`⭐️プロジェクトID⭐️`は、`.clasp.json`に記載されています。  

![プロパティサービスへのデータの登録](./docs/img/property-service.gif)  

登録するデータは`.env.example`に記載しています。  

[公式ドキュメント](https://developers.google.com/apps-script/guides/properties?hl=ja)  

## デプロイ

`cat ~/.clasprc.json`でGASの認証情報を取得します。  
これをGitHubのシークレットに`CLASP_RC_JSON`として登録します。  

`main`ブランチにpushすると、GitHub Actionsでデプロイされます。  
