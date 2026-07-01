FROM node:20-alpine

# Crie o diretório de dados para persistência
RUN mkdir -p /data

# Defina o diretório de trabalho no container
WORKDIR /app

# Copie os arquivos de dependências
COPY package*.json ./

# Instale apenas dependências de produção
RUN npm ci --only=production

# Copie o restante dos arquivos do projeto
COPY . .

# Expõe a porta que o servidor vai rodar
EXPOSE 3000

# Variáveis de ambiente padrão para Docker
ENV PORT=3000
ENV DATA_DIR=/data
ENV NODE_ENV=production

# Comando para iniciar o servidor
CMD ["node", "server.js"]
