#!/usr/bin/env python3
"""
Script per scaricare simboli, footprint e modelli 3D da LCSC usando easyeda2kicad
partendo da una BOM in formato CSV o Excel esportata da EasyEDA.
"""

import argparse
import subprocess
import pandas as pd
from typing import List
import os

def parse_arguments() -> argparse.Namespace:
    """
    Configura e parsa gli argomenti da riga di comando.
    """
    parser = argparse.ArgumentParser(
        description="Scarica simboli/footprint da LCSC con easyeda2kicad partendo da una BOM",
        epilog="Esempi:\n"
               "  ./bom_to_kicad.py --bom BOM.csv\n"
               "  ./bom_to_kicad.py --bom BOM.xlsx --sheet_name 'Elenco componenti'\n"
               "  ./bom_to_kicad.py --bom BOM.csv --column 'LCSC' --full"
    )
    parser.add_argument("--bom", help="Percorso al file BOM (CSV o Excel)")
    parser.add_argument("--delimiter", default="auto", help="Separatore per CSV (default: rilevamento automatico)")
    parser.add_argument("--sheet_name", default=None, help="Nome del foglio Excel (per file XLSX)")
    parser.add_argument("--column", default="Supplier Part", help="Nome colonna con codice LCSC (default: 'Supplier Part')")
    parser.add_argument("--full", action="store_true", help="Scarica simbolo, footprint e 3D")
    parser.add_argument("--symbol", action="store_true", help="Scarica solo il simbolo")
    parser.add_argument("--footprint", action="store_true", help="Scarica solo il footprint")
    parser.add_argument("--overwrite", action="store_true", help="Sovrascrivi i file esistenti")
    parser.add_argument("--3d", dest="model3d", action="store_true", help="Scarica solo il modello 3D") # Per compatibilità con easyeda2kicad
    parser.add_argument("--output", action="store", type=str, help="Imposta la directory di output (default: ~/libs/my_lib per i simboli e ~/libs/footprints per i footprint)")
    parser.add_argument("--single-component", default=True, type=bool, help="Consente di scaricare un singolo componente LCSC specificato da riga di comando")

    return parser.parse_args()

def detect_file_type(file_path: str) -> str:
    """Rileva se il file è CSV o Excel in base all'estensione."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        print("Rilevato file formato Excel")
        return 'excel'
    print("Rilevato file formato CSV")
    return 'csv'

def read_excel(file_path: str, sheet_name: str = None) -> pd.DataFrame:
    """Legge un file Excel con gestione degli errori."""
    try:
        if sheet_name:
            return pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        return pd.read_excel(file_path, dtype=str)
    except Exception as e:
        available_sheets = pd.ExcelFile(file_path).sheet_names
        raise ValueError(
            f"Errore lettura Excel: {str(e)}\n"
            f"Fogli disponibili: {', '.join(available_sheets)}"
        )

def read_csv(file_path: str, delimiter: str = "auto") -> pd.DataFrame:
    """Legge un file CSV con rilevamento automatico del delimitatore."""
    encodings = ['utf-8-sig', 'cp1252']
    
    if delimiter == "auto":
        # Prova a rilevare automaticamente il delimitatore
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    first_line = f.readline()
                
                if '\t' in first_line:
                    delimiter = '\t'
                elif ',' in first_line:
                    delimiter = ','
                elif ';' in first_line:
                    delimiter = ';'
                else:
                    delimiter = ','  # Default
                break
            except UnicodeDecodeError:
                continue
    
    # Prova a leggere con i diversi encoding
    for enc in encodings:
        try:
            return pd.read_csv(
                file_path,
                delimiter=delimiter,
                quotechar='"',
                encoding=enc,
                dtype=str
            )
        except UnicodeDecodeError:
            continue
    
    raise ValueError("Impossibile leggere il file con gli encoding supportati (UTF-8, Windows-1252)")

def extract_lcsc_parts(bom_path: str, delimiter: str, column_name: str, sheet_name: str = None, single_component: bool = False) -> List[str]:
    """
    Estrae i codici LCSC dalla BOM (CSV o Excel).
    """

    
    file_type = detect_file_type(bom_path)
    
    try:
        if file_type == 'excel':
            df = read_excel(bom_path, sheet_name)
        else:
            df = read_csv(bom_path, delimiter)
    except Exception as e:
        raise ValueError(f"Errore lettura file {bom_path}: {str(e)}")

    # Debug: mostra le colonne disponibili
    print("\n📋 Colonne disponibili nel BOM:")
    print(df.columns.tolist())

    if column_name not in df.columns:
        raise ValueError(
            f"Colonna '{column_name}' non trovata. Colonne disponibili: {', '.join(df.columns)}"
        )

    # Estrai codici LCSC, pulisci e rimuovi duplicati
    parts = (
        df[column_name]
        .dropna()
        .astype(str)
        .str.strip()
        .replace(r'^\s*$', pd.NA, regex=True)
        .dropna()
        .unique()
    )

    print(f"\n🔎 Trovati {len(parts)} codici LCSC unici nella colonna '{column_name}'")
    return sorted(parts)

def build_easyeda2kicad_args(args: argparse.Namespace) -> List[str]:
    """Costruisce gli argomenti per easyeda2kicad."""  
    
    cli_args = []

    if args.full:
        cli_args.append("--full")
    elif args.symbol:
        cli_args.append("--symbol")
    elif args.footprint:
        cli_args.append("--footprint")
    elif args.model3d:
        cli_args.append("--3d")
    else:
        cli_args.append("--full")
    
    if args.overwrite:
        cli_args.append("--overwrite")
    
    if args.output:
        output_dir = os.path.abspath(args.output)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        cli_args.append(f"--output={output_dir}")

    return cli_args if cli_args else ["--full"]

def download_component(part_number: str, easyeda_args: List[str]) -> bool:
    """Scarica un componente usando easyeda2kicad."""
    print(f"⬇️  Scaricando {part_number}...")
    cmd = ["easyeda2kicad"] + easyeda_args + [f"--lcsc_id={part_number}"]
    
    try:
        result = subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        print(f"✅ Successo per {part_number}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Fallito {part_number}: {e.stderr.strip()}")
        return False
    except FileNotFoundError:
        print("❌ easyeda2kicad non trovato. Installalo con: pip install easyeda2kicad")
        return False

def main():
    args = parse_arguments()
    
    try:
        parts = []
        
        if not args.bom and args.single_component:
            part_number = input("Inserisci il codice LCSC del componente da scaricare: ").strip()
            if not part_number:
                print("❌ Codice LCSC non fornito. Uscita.")
                return
            parts = [part_number]
            print(f"📦 Componente da scaricare: {part_number}")

        elif args.bom and not args.single_component:

            bom_path = None
            if args.bom:
                bom_path = os.path.abspath(args.bom)
                if not os.path.exists(bom_path):
                    raise FileNotFoundError(f"File BOM non trovato: {bom_path}")
            else:
                bom_path = str(input("Inserisci il percorso al file BOM (CSV o Excel): ").strip())
                print(f"📂 Percorso BOM: {bom_path}")
                if not bom_path:
                    print("❌ Percorso BOM non fornito. Uscita.")
                    return

            parts = extract_lcsc_parts(
                bom_path,
                args.delimiter,
                args.column,
                args.sheet_name
            )
        
        if not parts:
            print("⚠️ Nessun codice LCSC trovato nella BOM")
            return
        
        cli_args = build_easyeda2kicad_args(args)
        
        print(f"\n⚙️  Opzioni easyeda2kicad: {' '.join(cli_args)}")
        print(f"🔧 Componenti da processare: {len(parts)}\n")
        
        success_count = 0
        for part in parts:
            if download_component(part, cli_args):
                success_count += 1
        
        print(f"\n🎉 Riepilogo:")
        print(f"- Componenti totali: {len(parts)}")
        print(f"- Scaricati con successo: {success_count}")
        print(f"- Falliti: {len(parts) - success_count}")
        
    except Exception as e:
        print(f"\n❌ Errore: {str(e)}")
        exit(1)

if __name__ == "__main__":
    main()