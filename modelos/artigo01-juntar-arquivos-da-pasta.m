// Artigo 1 - juntar arquivos de uma pasta (consulta em branco > Editor Avançado).
// Troque Pasta pelo seu caminho, COM a barra final: sem ela o filtro DaPasta devolve 0 linhas.
// Resultado executado no Excel (excelcli), pasta com Filial_Norte/Sul/Centro.xlsx,
// um arquivo de trava ~$Filial_Sul.xlsx e uma subpasta Antigos: 30 linhas, colunas Name | Data | Item | Valor.
// Sem o filtro DaPasta: 40 linhas (a subpasta entra). Sem o filtro do ~$: DataFormat.Error.
// Supõe uma aba chamada "Vendas" com as colunas Data, Item e Valor. Fonte: site/testes/m/a1-final.m, a1a4-saida.txt.
let
    Pasta = "C:\SiteTeste\Filiais\",
    Arquivos = Folder.Files(Pasta),
    DaPasta = Table.SelectRows(Arquivos, each [Folder Path] = Pasta),
    SoExcel = Table.SelectRows(
        DaPasta,
        each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")
    ),
    Lidos = Table.AddColumn(
        SoExcel,
        "Dados",
        each Excel.Workbook([Content], true){[Item = "Vendas", Kind = "Sheet"]}[Data]
    ),
    Mantidas = Table.SelectColumns(Lidos, {"Name", "Dados"}),
    Expandido = Table.ExpandTableColumn(Mantidas, "Dados", {"Data", "Item", "Valor"}),
    Tipado = Table.TransformColumnTypes(Expandido, {{"Data", type date}, {"Valor", type number}})
in
    Tipado
