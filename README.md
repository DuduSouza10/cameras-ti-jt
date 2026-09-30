# Acompanhamento de Bases — Railway

Dashboard web para acompanhar status das bases, dias de gravação, quantidade de câmeras e câmeras off-line, preservando um histórico de cada atualização.

## O que já está pronto
- Importação manual de XLSX.
- Snapshot histórico a cada importação.
- Comparação automática com o snapshot anterior.
- Indicadores, filtros, alertas e gráficos.
- Exportação de relatório XLSX.
- Endpoint `POST /api/ingest` preparado para a futura integração com Feishu.
- PostgreSQL automático quando existir `DATABASE_URL`; fallback local em SQLite.

## Deploy no Railway
1. Suba esta pasta para um repositório GitHub.
2. Crie um novo projeto no Railway a partir do repositório.
3. Adicione um serviço PostgreSQL no mesmo projeto.
4. Garanta que a variável `DATABASE_URL` esteja disponível no serviço do dashboard.
5. Faça o deploy. O `railway.toml` já define o start e o healthcheck.

> Sem PostgreSQL, o app usa SQLite. Em produção no Railway, use PostgreSQL para o histórico permanecer persistente entre deploys.

## Integração futura com Feishu
Envie para `POST /api/ingest` um JSON como:

```json
{
  "source": "feishu",
  "source_date": "30/09/2026",
  "rows": [
    {
      "base": "AUI-MG",
      "recording_days": "14 DIAS",
      "camera_count": "5 CAM",
      "offline_count": "0 CAM",
      "status": "EM FUNCIONAMENTO"
    }
  ]
}
```

Cada envio cria um novo snapshot. Na integração definitiva, a autenticação do endpoint deve ser protegida por token ou assinatura do Feishu.
