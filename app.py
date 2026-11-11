def prepare_and_convert_excel(uploaded_file):
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, uploaded_file.name)
        
        # 1. Guardar original
        with open(input_path, "wb") as f:
            f.write(uploaded_file.read())
            
        numero_informe_detectado = None
            
        # 2. Leer el Excel y limpiar hojas vacías
        if input_path.endswith('.xlsx'):
            wb = openpyxl.load_workbook(input_path)
            hojas_a_borrar = []
            
            for nombre_hoja in wb.sheetnames:
                ws = wb[nombre_hoja]
                hoja_tiene_datos = False
                
                # Buscar en las primeras 15 filas y 5 columnas el Número de Informe
                for row in range(1, 15):
                    for col in range(1, 6):
                        valor_celda = str(ws.cell(row=row, column=col).value or "").strip()
                        if "R-LP-" in valor_celda or "INFORME N" in valor_celda.upper():
                            if len(valor_celda) > 10: # Para asegurar que no es solo el título vacío
                                hoja_tiene_datos = True
                                numero_informe_detectado = valor_celda.replace("INFORME N°", "").replace(":", "").strip()
                                break
                    if hoja_tiene_datos:
                        break
                
                if not hoja_tiene_datos:
                    hojas_a_borrar.append(nombre_hoja)
            
            # Borrar las hojas que son plantillas vacías (dejamos solo la que tiene el informe)
            if len(hojas_a_borrar) < len(wb.sheetnames): # Asegurar no borrar todo
                for hoja in hojas_a_borrar:
                    del wb[hoja]
            
            # Forzar escala a 1 página
            for ws in wb.worksheets:
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0
                
            wb.save(input_path)

        # 3. Exportar con LibreOffice
        comando = [
            "libreoffice", "--headless", "--nologo", "--nofirststartwizard", 
            "--convert-to", "pdf", input_path, "--outdir", tmpdir
        ]
        subprocess.run(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # 4. Leer PDF (y si encontramos el número de informe, sugerir ese nombre)
        pdf_filename = uploaded_file.name.rsplit('.', 1)[0] + ".pdf"
        pdf_path = os.path.join(tmpdir, pdf_filename)
        
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
            
        nombre_final = f"{numero_informe_detectado}.pdf" if numero_informe_detectado else pdf_filename
        return pdf_bytes, nombre_final
