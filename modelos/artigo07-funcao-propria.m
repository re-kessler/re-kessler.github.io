// Artigo 7 - função própria e uma chamada por filial. São DUAS consultas.
// Executado no Excel: Todas = 30 linhas (10 por filial), colunas Data | Item | Valor | Filial,
// idêntica à versão numa consulta só. Se a coluna Filial já existir na planilha de origem:
// Expression.Error "O campo 'Filial' já existe no registro". Fonte: site/testes/m/a7-*.m, a7-saida.txt.

// ---- Consulta 1, nome: fnFilial (a função; troque o caminho dos arquivos pelo seu) ----
(Filial as text) as table =>
let
    Livro = Excel.Workbook(File.Contents("C:\SiteTeste\Filiais\Filial_" & Filial & ".xlsx"), true),
    Vendas = Livro{[Item = "Vendas", Kind = "Sheet"]}[Data],
    Tipado = Table.TransformColumnTypes(Vendas, {{"Data", type date}, {"Valor", type number}}),
    ComFilial = Table.AddColumn(Tipado, "Filial", each Filial, type text)
in
    ComFilial

// ---- Consulta 2, nome: Todas ----
let
    Filiais = {"Norte", "Sul", "Centro"},
    Tabelas = List.Transform(Filiais, fnFilial),
    Todas = Table.Combine(Tabelas)
in
    Todas
