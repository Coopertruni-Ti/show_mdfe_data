import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io


def extrair_ctes(infCTe):
    ctes_info = []
    for cte in infCTe:
        chaveCTe_element = cte.find(
            "{http://www.portalfiscal.inf.br/mdfe}chCTe")
        if chaveCTe_element is None:
            continue

        chaveCTe = chaveCTe_element.text or ""
        serie = chaveCTe[-21:-19].lstrip('0')
        codigoExtraido = chaveCTe[-19:-10].lstrip('0')

        try:
            codigoInt = int(codigoExtraido)
        except ValueError:
            codigoInt = 0

        ctes_info.append({
            "codigoInt": codigoInt,
            "serie": serie,
            "chaveCTe": chaveCTe,
            "codigoExtraido": codigoExtraido,
        })
    return ctes_info


def filtro_placas(placa1, placa2, include_placas, ignore_placas):
    placa1_up = placa1.upper()
    placa2_up = placa2.upper()
    include_placas_up = [p.upper() for p in include_placas]
    ignore_placas_up = [p.upper() for p in ignore_placas]

    if placa1_up in ignore_placas_up or placa2_up in ignore_placas_up:
        return False

    if include_placas_up:
        return (placa1_up in include_placas_up) or (placa2_up in include_placas_up)

    return True


def parse_mdfes(arquivos):
    mdfes_por_ctes = {}

    for file in arquivos:
        try:
            xml_content = file.read()
            file.seek(0)
            tree = ET.parse(io.BytesIO(xml_content))
            root = tree.getroot()

            ide = root.find(".//{http://www.portalfiscal.inf.br/mdfe}ide")
            if ide is None:
                continue

            nMDF_element = ide.find(
                "{http://www.portalfiscal.inf.br/mdfe}nMDF")
            if nMDF_element is None or not nMDF_element.text:
                continue
            nMDF = int(nMDF_element.text)

            dhEmi_element = ide.find(
                "{http://www.portalfiscal.inf.br/mdfe}dhEmi")
            dhEmi = dhEmi_element.text if dhEmi_element is not None else ""
            dataEmissao = ""
            if len(dhEmi) >= 10:
                dataEmissao = f"{dhEmi[8:10]}/{dhEmi[5:7]}/{dhEmi[:4]}"

            infModal = root.find(
                ".//{http://www.portalfiscal.inf.br/mdfe}infModal")
            veicTracao = infModal.find(
                ".//{http://www.portalfiscal.inf.br/mdfe}veicTracao") if infModal is not None else None
            veicReboque = infModal.find(
                ".//{http://www.portalfiscal.inf.br/mdfe}veicReboque") if infModal is not None else None

            placa1 = ""
            if veicTracao is not None:
                placa1_element = veicTracao.find(
                    "{http://www.portalfiscal.inf.br/mdfe}placa")
                placa1 = placa1_element.text if placa1_element is not None else ""

            placa2 = ""
            if veicReboque is not None:
                placa2_element = veicReboque.find(
                    "{http://www.portalfiscal.inf.br/mdfe}placa")
                placa2 = placa2_element.text if placa2_element is not None else ""

            infMunCarrega = ide.find(
                "{http://www.portalfiscal.inf.br/mdfe}infMunCarrega")
            localCarregamento = ""
            if infMunCarrega is not None:
                xMunCarrega_element = infMunCarrega.find(
                    "{http://www.portalfiscal.inf.br/mdfe}xMunCarrega")
                localCarregamento = xMunCarrega_element.text if xMunCarrega_element is not None else ""

            condutores = root.findall(
                ".//{http://www.portalfiscal.inf.br/mdfe}condutor")
            if condutores:
                condutor = condutores[0]
                nome_motorista_element = condutor.find(
                    "{http://www.portalfiscal.inf.br/mdfe}xNome")
                cpf_motorista_element = condutor.find(
                    "{http://www.portalfiscal.inf.br/mdfe}CPF")
                nome_motorista = nome_motorista_element.text if nome_motorista_element is not None else "N/A"
                cpf_motorista = cpf_motorista_element.text if cpf_motorista_element is not None else "N/A"
            else:
                nome_motorista = "N/A"
                cpf_motorista = "N/A"

            infCTe_list = root.findall(
                ".//{http://www.portalfiscal.inf.br/mdfe}infCTe")
            ctes_info = extrair_ctes(infCTe_list)

            if not ctes_info:
                continue

            codigosCTe = sorted([info["codigoExtraido"] for info in ctes_info])
            chavesCTe = [info["chaveCTe"] for info in ctes_info]
            series = sorted(list({info["serie"] for info in ctes_info}))

            ctes_set = frozenset(codigosCTe)

            # Mantém o MDF-e de maior nMDF em caso de duplicidade
            if ctes_set in mdfes_por_ctes:
                if nMDF > mdfes_por_ctes[ctes_set]["nMDF"]:
                    mdfes_por_ctes[ctes_set] = {
                        "nMDF": nMDF,
                        "Data Emissão": dataEmissao,
                        "Placa 1": placa1,
                        "Placa 2": placa2,
                        "Local Carregamento": localCarregamento,
                        "Nome Motorista": nome_motorista,
                        "CPF Motorista": cpf_motorista,
                        "Códigos CTe": codigosCTe,
                        "Chaves CTe": chavesCTe,
                        "Séries": series,
                    }
            else:
                mdfes_por_ctes[ctes_set] = {
                    "nMDF": nMDF,
                    "Data Emissão": dataEmissao,
                    "Placa 1": placa1,
                    "Placa 2": placa2,
                    "Local Carregamento": localCarregamento,
                    "Nome Motorista": nome_motorista,
                    "CPF Motorista": cpf_motorista,
                    "Códigos CTe": codigosCTe,
                    "Chaves CTe": chavesCTe,
                    "Séries": series,
                }

        except ET.ParseError as e:
            print(f"Erro ao analisar o arquivo: {e}")
        except Exception as ex:
            print(f"Erro inesperado: {ex}")

    if not mdfes_por_ctes:
        return pd.DataFrame()

    linhas = []
    for data in mdfes_por_ctes.values():
        linhas.append([
            data["nMDF"],
            data["Data Emissão"],
            data["Placa 1"],
            data["Placa 2"],
            data["Local Carregamento"],
            data["Nome Motorista"],
            data["CPF Motorista"],
            data["Códigos CTe"],
            data["Chaves CTe"],
            data["Séries"],
        ])

    df = pd.DataFrame(
        linhas,
        columns=[
            "nMDF",
            "Data Emissão",
            "Placa 1",
            "Placa 2",
            "Local Carregamento",
            "Nome Motorista",
            "CPF Motorista",
            "Códigos CTe",
            "Chaves CTe",
            "Séries",
        ]
    )
    return df


def filtrar_dataframe(df,
                      series_selecionadas,
                      codigo_inicial,
                      codigo_final,
                      ctes_list,
                      ignore_ctes,
                      include_placas,
                      ignore_placas):
    df_filtrado = df.copy()

    # 1) Séries
    if series_selecionadas:
        df_filtrado = df_filtrado[df_filtrado["Séries"].apply(
            lambda sers: any(serie in sers for serie in series_selecionadas)
        )]

    # 2) Placas
    mask_placas = []
    for _, row in df_filtrado.iterrows():
        placa1 = row["Placa 1"]
        placa2 = row["Placa 2"]
        mask_placas.append(filtro_placas(
            placa1, placa2, include_placas, ignore_placas))
    df_filtrado = df_filtrado[mask_placas]

    # 3) CT-es
    linhas_validas = []
    for _, row in df_filtrado.iterrows():
        codigos_cte_str = row["Códigos CTe"]
        codigos_int = []
        for cte_str in codigos_cte_str:
            try:
                codigos_int.append(int(cte_str))
            except:
                pass

        # a) Faixa
        if not all(codigo_inicial <= c <= codigo_final for c in codigos_int):
            linhas_validas.append(False)
            continue

        # b) ctes_list
        if ctes_list:
            if not any(c in ctes_list for c in codigos_int):
                linhas_validas.append(False)
                continue

        # c) ignore_ctes
        if ignore_ctes:
            if any(c in ignore_ctes for c in codigos_int):
                linhas_validas.append(False)
                continue

        linhas_validas.append(True)

    df_filtrado = df_filtrado[linhas_validas]

    return df_filtrado


def dataframe_to_excel(df):
    """
    Converte um DataFrame em bytes para Excel.
    Retorna o conteúdo binário pronto para uso em um st.download_button.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
    return output.getvalue()


def main():
    st.title("Processamento de MDF-e")

    if "df_completo" not in st.session_state:
        st.session_state.df_completo = pd.DataFrame()

    st.write(
        "**1) Faça o upload dos arquivos XML de MDF-e** (você pode selecionar vários):")
    uploaded_files = st.file_uploader(
        "Selecione os XMLs de MDF-e",
        type=["xml"],
        accept_multiple_files=True
    )

    if st.button("Carregar MDF-es"):
        if uploaded_files:
            df_completo = parse_mdfes(uploaded_files)
            st.session_state.df_completo = df_completo

            if df_completo.empty:
                st.warning(
                    "Não foi possível extrair dados válidos desses XMLs.")
            else:
                st.success(
                    "Arquivos carregados com sucesso! Selecione os filtros abaixo.")
        else:
            st.warning("Nenhum arquivo selecionado.")

    # Exibe os filtros somente se já tiver MDF-es carregados
    if not st.session_state.df_completo.empty:
        df_completo = st.session_state.df_completo

        st.write("---")
        st.subheader("Filtros")

        todas_series = sorted(
            list(
                {serie for lista_s in df_completo["Séries"] for serie in lista_s})
        )

        series_selecionadas = st.multiselect(
            "Filtrar por Série (deixe vazio para incluir todas)",
            options=todas_series
        )

        codigo_inicial = st.number_input("Código Inicial (CT-e)", value=0)
        codigo_final = st.number_input("Código Final (CT-e)", value=999999999)

        ctes_list_input = st.text_input(
            "CT-es a Incluir (separados por vírgula) - deixe vazio para não filtrar")
        if ctes_list_input.strip():
            ctes_list = [int(x.strip()) for x in ctes_list_input.split(
                ",") if x.strip().isdigit()]
        else:
            ctes_list = []

        ignore_ctes_input = st.text_input(
            "CT-es a Ignorar (separados por vírgula)")
        if ignore_ctes_input.strip():
            ignore_ctes = [int(x.strip()) for x in ignore_ctes_input.split(
                ",") if x.strip().isdigit()]
        else:
            ignore_ctes = []

        include_placas_input = st.text_input(
            "Placas a Incluir (separadas por vírgula) - vazio = todas permitidas")
        include_placas = [p.strip().upper() for p in include_placas_input.split(
            ",") if p.strip()] if include_placas_input else []

        ignore_placas_input = st.text_input(
            "Placas a Ignorar (separadas por vírgula)")
        ignore_placas = [p.strip().upper() for p in ignore_placas_input.split(
            ",") if p.strip()] if ignore_placas_input else []

        if st.button("Aplicar Filtros"):
            df_filtrado = filtrar_dataframe(
                df=df_completo,
                series_selecionadas=series_selecionadas,
                codigo_inicial=codigo_inicial,
                codigo_final=codigo_final,
                ctes_list=ctes_list,
                ignore_ctes=ignore_ctes,
                include_placas=include_placas,
                ignore_placas=ignore_placas
            )

            if df_filtrado.empty:
                st.warning(
                    "Nenhuma viagem atende aos critérios especificados.")
            else:
                st.subheader("Resultado do Processamento:")
                st.dataframe(df_filtrado)

                # ---- Botão para baixar em Excel ----
                excel_data = dataframe_to_excel(df_filtrado)
                st.download_button(
                    label="Baixar em Excel",
                    data=excel_data,
                    file_name="resultado_mdfes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
