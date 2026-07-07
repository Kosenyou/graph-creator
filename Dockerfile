# Node.js 20 LTS の軽量イメージを使用
FROM node:20-slim

# アプリケーションディレクトリを作成する
WORKDIR /usr/src/app

# パッケージ定義ファイルをコピー
COPY package*.json ./

# 本番環境用の依存関係のみをインストール
RUN npm ci --only=production

# ソースコードをコピーする
COPY . .

# サービスがリッスンするポートを指定（Cloud Runのデフォルト環境変数 PORT に対応）
EXPOSE 8080

# サーバーを起動
CMD [ "npm", "start" ]
