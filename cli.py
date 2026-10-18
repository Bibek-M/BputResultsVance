"""
BPUT Results Scraper & Excel Exporter CLI
Terminal interface with rich styling, interactive prompts, and command-line flags.
"""

import os
import sys
import argparse
from typing import Optional

# Ensure UTF-8 output on Windows terminals to prevent charmap encoding errors
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.prompt import Prompt, Confirm

from bput_scraper import BPUTScraper
from excel_exporter import BPUTExcelExporter

console = Console()



def display_banner():
    banner_text = """[bold cyan]╔═══════════════════════════════════════════════════════════════════════════╗
║          [bold yellow]BIJU PATNAIK UNIVERSITY OF TECHNOLOGY (BPUT) RESULTS[/bold yellow]          ║
║               [bold white]Automated Academic Record Scraper & Excel Generator[/bold white]         ║
╚═══════════════════════════════════════════════════════════════════════════╝[/bold cyan]
[dim italic]Portal Source: https://results.bput.ac.in/ | Fast Multi-threaded Academic Scraper[/dim italic]"""
    console.print(Panel(banner_text, border_style="bright_blue", expand=False))


def display_results_in_terminal(data: dict):
    """Prints a structured summary table directly in the terminal using Rich."""
    info = data.get("student_info", {})
    summary = data.get("summary", {})
    semesters = data.get("semesters", [])

    # 1. Student Profile Table
    profile_tbl = Table(title="🎓 Student Academic Profile", border_style="bright_blue", show_header=False)
    profile_tbl.add_column("Key", style="bold cyan", width=22)
    profile_tbl.add_column("Value", style="bold white")

    profile_tbl.add_row("Student Name", info.get("studentName", "N/A"))
    profile_tbl.add_row("Registration Number", info.get("rollNo", "N/A"))
    profile_tbl.add_row("Course / Degree", info.get("courseName", "N/A"))
    profile_tbl.add_row("Branch / Department", info.get("branchName", "N/A"))
    profile_tbl.add_row("College / Institute", info.get("collegeName", "N/A"))
    profile_tbl.add_row("Admission Batch", str(info.get("batch", "N/A")))
    profile_tbl.add_row("Academic Standing", f"[bold green]{summary.get('academic_standing')}[/bold green]" if "CLEARED" in str(summary.get("academic_standing")) else f"[bold red]{summary.get('academic_standing')}[/bold red]")

    console.print()
    console.print(profile_tbl)

    # 2. Key Performance Indicators
    kpi_tbl = Table(title="📊 Cumulative Metrics", border_style="green", show_lines=True)
    kpi_tbl.add_column("Cumulative CGPA", justify="center", style="bold yellow")
    kpi_tbl.add_column("Total Credits", justify="center", style="bold green")
    kpi_tbl.add_column("Total Grade Points", justify="center", style="bold cyan")
    kpi_tbl.add_column("Semesters Scraped", justify="center", style="bold white")
    kpi_tbl.add_column("Unique Subjects", justify="center", style="bold magenta")
    kpi_tbl.add_column("Active Backlogs", justify="center", style="bold red" if summary.get("active_backlog_count", 0) > 0 else "bold green")

    kpi_tbl.add_row(
        f"{summary.get('cgpa', 0.0):.2f}",
        str(summary.get("total_credits", 0)),
        str(summary.get("total_grade_points", 0.0)),
        str(summary.get("total_semesters_found", 0)),
        str(summary.get("total_unique_subjects", 0)),
        str(summary.get("active_backlog_count", 0))
    )
    console.print()
    console.print(kpi_tbl)

    # 3. Semester Breakdown Table
    if semesters:
        sem_tbl = Table(title="📑 Semester-wise Examination Record", border_style="cyan", show_lines=True)
        sem_tbl.add_column("Sem", justify="center", style="bold cyan")
        sem_tbl.add_column("Session", justify="center", style="white")
        sem_tbl.add_column("Type", justify="center", style="dim")
        sem_tbl.add_column("Subjects", justify="center", style="magenta")
        sem_tbl.add_column("Credits", justify="center", style="green")
        sem_tbl.add_column("Points", justify="right", style="cyan")
        sem_tbl.add_column("SGPA", justify="center", style="bold yellow")
        sem_tbl.add_column("Status", justify="center")

        for sem in semesters:
            status_str = f"[bold green]{sem['status']}[/bold green]" if "PASS" in sem['status'] else f"[bold red]{sem['status']}[/bold red]"
            sgpa_disp = f"{sem['sgpa']:.2f}" if sem.get("sgpa") else "N/A"
            sem_tbl.add_row(
                str(sem.get("semester", "")),
                str(sem.get("session", "")),
                str(sem.get("exam_type", "")),
                str(len(sem.get("grades", []))),
                str(sem.get("credits", 0)),
                str(round(sem.get("points", 0.0), 2)),
                sgpa_disp,
                status_str
            )
        console.print()
        console.print(sem_tbl)


def run_cli():
    parser = argparse.ArgumentParser(description="Scrape BPUT student exam results and generate styled Excel spreadsheets.")
    parser.add_argument("-r", "--reg", "--roll", dest="roll_no", type=str, help="Student 10-digit Registration / Roll Number")
    parser.add_argument("-d", "--dob", dest="dob", type=str, default="2000-01-01", help="Date of Birth in YYYY-MM-DD format (default: 2000-01-01)")

    parser.add_argument("-s", "--session", dest="session", type=str, default=None, help="Specific exam session name (e.g. 'Even-(2023-24)')")
    parser.add_argument("-o", "--output", dest="output", type=str, default=None, help="Custom output Excel filename / path (.xlsx)")
    parser.add_argument("--open", dest="auto_open", action="store_true", help="Automatically launch the generated Excel file upon completion")
    parser.add_argument("--all", dest="all_sessions", action="store_true", help="Scrape all available sessions without interactive prompt")

    args = parser.parse_args()

    display_banner()

    scraper = BPUTScraper()

    # 1. Registration Number
    roll_no = args.roll_no
    if not roll_no:
        while True:
            roll_no = Prompt.ask("\n[bold yellow]👉 Enter Student Registration Number[/bold yellow]").strip()
            if not roll_no:
                console.print("[red]Registration number cannot be empty![/red]")
                continue
            if len(roll_no) != 10:
                console.print("[dim yellow]Note: Standard BPUT registration numbers are 10 digits (e.g., 2101289123).[/dim yellow]")
            break

    # 2. Date of Birth (defaulted automatically to 2000-01-01)
    dob = args.dob or BPUTScraper.DEFAULT_DOB



    # 3. Available Sessions
    with console.status("[bold green]Connecting to BPUT portal & retrieving exam sessions...[/bold green]", spinner="dots"):
        all_sessions = scraper.get_available_sessions()

    # 4. Session Scope Selection
    sessions_to_scrape = None
    if args.session:
        # User specified session via flag
        matched = [s for s in all_sessions if args.session.lower() in s.lower()]
        if matched:
            sessions_to_scrape = [matched[0]]
            console.print(f"[green]Selected exam session:[/green] [bold]{matched[0]}[/bold]")
        else:
            sessions_to_scrape = [args.session]
            console.print(f"[yellow]Using specified session:[/yellow] [bold]{args.session}[/bold]")
    elif not args.all_sessions:
        console.print(f"\n[bold cyan]Found {len(all_sessions)} exam sessions on BPUT portal.[/bold cyan]")
        console.print("  [1] [bold green]Full Degree History[/bold green] (Scans all sessions — Recommended)")
        console.print("  [2] [bold yellow]Select a Specific Session[/bold yellow]")

        choice = Prompt.ask("[bold yellow]👉 Select scraping mode[/bold yellow]", choices=["1", "2"], default="1")
        if choice == "2":
            console.print("\n[bold]Available Sessions:[/bold]")
            for idx, sess in enumerate(all_sessions, start=1):
                console.print(f"  [{idx:2d}] {sess}")
            sess_idx = Prompt.ask(f"[bold yellow]👉 Enter session number (1-{len(all_sessions)})[/bold yellow]", default="1")
            try:
                sel = all_sessions[int(sess_idx) - 1]
                sessions_to_scrape = [sel]
                console.print(f"[green]Selected session:[/green] [bold]{sel}[/bold]")
            except Exception:
                console.print("[red]Invalid selection! Defaulting to all sessions.[/red]")
                sessions_to_scrape = None

    # 5. Fetch Full Results with Progress Display
    console.print(f"\n[bold bright_blue]🚀 Starting scraping for Reg No: [yellow]{roll_no}[/yellow][/bold bright_blue]")

    data = None
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        console=console
    ) as progress:
        task = progress.add_task("[cyan]Scanning exam sessions...", total=len(sessions_to_scrape or all_sessions))

        def on_progress(event, payload):
            if event == "session_checked":
                progress.advance(task, 1)
            elif event == "grades_start":
                progress.update(task, description=f"[green]Extracting subject grades for {payload} semesters...")

        data = scraper.fetch_full_results(
            roll_no=roll_no,
            dob=dob,
            sessions_to_check=sessions_to_scrape,
            progress_callback=on_progress
        )

    if not data or not data.get("semesters"):
        console.print(f"\n[bold red]❌ No published results found for Registration Number: {roll_no}[/bold red]")
        console.print("[dim]Please check if the registration number is correct or if results for this student are published on https://results.bput.ac.in/.[/dim]")
        sys.exit(0)

    # 6. Display Terminal Summary
    display_results_in_terminal(data)

    # 7. Generate Excel File
    console.print("\n[bold cyan]📊 Generating styled Excel spreadsheet...[/bold cyan]")
    exporter = BPUTExcelExporter(data)
    out_file = exporter.export(args.output)
    file_size_kb = os.path.getsize(out_file) / 1024

    console.print(Panel(
        f"[bold green]✅ Excel file created successfully![/bold green]\n\n"
        f"📄 [bold white]File:[/bold white] [cyan]{os.path.basename(out_file)}[/cyan]\n"
        f"📂 [bold white]Path:[/bold white] [link=file:///{out_file.replace(os.sep, '/')}]file:///{out_file.replace(os.sep, '/')}[/link]\n"
        f"📏 [bold white]Size:[/bold white] {file_size_kb:.1f} KB\n"
        f"📑 [bold white]Sheets:[/bold white] 'Academic Overview', 'Consolidated Grade Sheet', 'All Subjects Master List'",
        title="[bold green]Export Complete[/bold green]",
        border_style="green"
    ))

    # 8. Open file prompt
    should_open = args.auto_open
    if not should_open and sys.stdin.isatty():
        should_open = Confirm.ask("[bold yellow]👉 Do you want to open the Excel spreadsheet now?[/bold yellow]", default=True)

    if should_open:
        try:
            os.startfile(out_file)
            console.print("[green]Opening Excel file...[/green]")
        except Exception as e:
            console.print(f"[dim red]Could not launch file automatically: {e}[/dim red]")


if __name__ == "__main__":
    run_cli()
