# ✅ PROJETO FINALIZADO — Caloi TPM v2.2

**Status:** EM PRODUÇÃO — NÃO ALTERAR  
**Data de finalização:** Maio 2025  
**URL Railway:** https://manutencao-tpm.up.railway.app  
**Repositório GitHub:** https://github.com/rcampos2000/caloi-tpm  

## ⚠️ Aviso Importante

Este projeto está em uso comercial ativo na Caloi Norte SA.  
**Nenhum arquivo deve ser modificado sem autorização.**

Qualquer nova funcionalidade deve ser desenvolvida no projeto **SGM_Caloi**,  
que é o sistema de próxima geração baseado neste projeto.

## Referências permitidas

O projeto SGM_Caloi pode **copiar e adaptar** os seguintes padrões deste projeto:
- Logo Caloi e marca d'água (`static/img/`)
- Ícones PWA (`static/icons/`)
- Padrão de autocomplete de TAG
- Reconhecimento de voz (Web Speech API)
- Preenchimento automático de setor e técnico
- Scanner de código de barras (html5-qrcode)
- Admin auto-save com sessionStorage
- Materiais com campo livre (checkboxes + chips)

## Stack técnica

- Flask 3.x + openpyxl + gunicorn + APScheduler
- Deploy: Railway (volume persistente `/data`)
- 216 equipamentos cadastrados
- Versão: **v2.2**
