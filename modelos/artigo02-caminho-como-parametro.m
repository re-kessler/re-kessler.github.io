// Artigo 2 - o caminho da pasta como parâmetro. São DUAS consultas: crie uma em branco
// para cada bloco e dê o nome indicado na primeira linha de cada um.
// Executado: com o parâmetro original, sem erro e 5 linhas (sem filtro de pasta);
// apontando para uma pasta que não existe, DataSource.NotFound. Fonte: site/testes/m/a2-*.m, a1a4-saida.txt.

// ---- Consulta 1, nome: CaminhoPasta (parâmetro; troque o texto pelo seu caminho) ----
"C:\SiteTeste\Filiais\" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]

// ---- Consulta 2, nome: Nomes ----
let
    Arquivos = Folder.Files(CaminhoPasta),
    SoExcel = Table.SelectRows(Arquivos, each Text.Lower([Extension]) = ".xlsx"),
    Nomes = Table.SelectColumns(SoExcel, {"Name", "Folder Path"})
in
    Nomes
