# vieira-silva-crm
CRM da Empresa Vieira e Silva

## Módulos
- **Gestão Comercial** — `index.html` (propostas, clientes, serviços, NFs, alertas).
- **Gestão e Acompanhamento de Obras** e **RDO** — `obras/gestao-obras.js`, carregado pelo `index.html`.
- **Mapa de andamento** por pavimento — `obras/<obra>/<bloco>.html` + `obras/_ponte-firebase.js`.

## Gestão de Obras — como funciona
| Onde | O quê |
|---|---|
| `obras/<obra>` | cadastro (nome, endereço, responsável, datas, situação) e equipe `membros {uid: gestor \| colaborador \| consulta}` |
| `obras/<obra>/itens` | EAP: etapa › subetapa › tarefa (avanço manual em % ou por produção) |
| `obras/<obra>/codigos/<código>` | trava que impede código repetido na mesma obra |
| `obras/<obra>/rdos/<AAAA-MM-DD>` | RDO (um por dia): rascunho → enviado → aprovado / cancelado |
| `obras/<obra>/lancamentos` | produção lançada — só inclusão. Id fixo `<rdo>__v<versão>__<tarefa>`: a mesma produção não entra duas vezes. Correção = estorno `__estorno` |
| `obras/<obra>/historico`, `serie` | auditoria (autor, data, antes/depois) e curva de evolução — só inclusão |
| `obras/<obra>/arquivos` | fotos (reduzidas para 1600 px) e anexos até 700 KB, guardados no Firestore |

Perfis: **admin** global = Administrador em todas as obras. **operador** recebe o papel definido na equipe de cada obra. **visualizador** só pode ser Consulta.

Avanço: tarefa por produção = executado ÷ previsto × 100 (previsto 0 → sem cálculo; acima de 100% é sinalizado e a quantidade real continua visível). O painel separa **% de tarefas concluídas** de **avanço físico ponderado**, que só aparece quando todas as tarefas têm peso; sem pesos, mostra “média simples” com o critério escrito.

## Publicar
1. **Regras do Firestore** — copie `firestore.rules` para *Console do Firebase › Firestore Database › Regras* e publique. O arquivo substitui as regras inteiras (já inclui Comercial, mapa e Gestão de Obras). Todo usuário precisa ter documento em `usuarios/<uid>` com `role`.
2. Faça o merge para `main` (GitHub Pages publica sozinho).
3. Entre como admin › Gestão de Obras › **Cadastrar agora** (traz a Clínica Ita Vilela do `obras.json` para o banco, sem apagar o mapa) e confira a equipe.
4. Opcional: **Criar obra de demonstração** (dados fictícios, marcada como DEMO; pode ser apagada em Cadastro e acesso).
