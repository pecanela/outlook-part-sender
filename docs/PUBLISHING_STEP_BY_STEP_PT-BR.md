# Publicação no GitHub — passo a passo para iniciantes

Use o pacote do **repositório** para preencher a branch `main`.
Os ZIPs executáveis e históricos devem ficar em **GitHub Releases**, não dentro de `main`.

## Estrutura esperada na raiz

A página principal deve mostrar diretamente:

`README.md`, `README.pt-BR.md`, `LICENSE`, `CHANGELOG.md`, `SECURITY.md`,
`CONTRIBUTING.md`, `VERSION`, `.github/`, `docs/`, `launcher/`,
`release-notes/` e `src/`.

Não envie uma pasta externa adicional chamada `outlook-part-sender` para dentro
de um repositório que já tenha esse nome.

## Release principal

- Tag: `v3.1.0`
- Título: `Outlook Part Sender v3.1.0 — Portable Corporate`
- Asset: `Outlook_Part_Sender_v3.1.0_PORTABLE_CORPORATE.zip`
- Estado: Latest
- Pre-release: não

Use `release-notes/v3.1.0.md` como descrição.

## Histórico recomendado

Para uma página mais limpa, publique inicialmente somente:

- `v1.0.0`
- `v1.4.0`
- `v2.0.2`
- `v3.0.0`
- `v3.1.0`

As demais versões permanecem documentadas em `CHANGELOG.md`.

## Verificação antes de tornar público

- não publicar documentos reais de clientes;
- não publicar endereços de e-mail reais;
- conferir screenshots;
- confirmar que você tem direito de publicar o código;
- baixar a release v3.1.0 após publicar e fazer um smoke test;
- manter Rascunhos como fluxo recomendado.
