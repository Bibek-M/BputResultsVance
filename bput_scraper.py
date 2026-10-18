"""
BPUT Results Scraper Engine
Fetches student examination results, personal profile, and semester-by-semester grades
from Biju Patnaik University of Technology (BPUT) portal (https://results.bput.ac.in/).
"""

import re
import time
from typing import Dict, List, Optional, Tuple, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests


class BPUTScraper:
    """Scraper client for BPUT exam results."""

    BASE_URL = "https://results.bput.ac.in"
    DEFAULT_DOB = "2000-01-01"  # BPUT portal accepts any valid date format

    # Default fallback list of sessions in case live extraction fails
    FALLBACK_SESSIONS = [
        "Supplementary 2025-26", "Even-(2025-26)", "Odd-(2025-26)",
        "Supplementary 2024-25", "Even-(2024-25)", "Odd-(2024-25)",
        "Supplementary 2023-24", "Even-(2023-24)", "Odd-(2023-24)",
        "Supplementary 2022-23", "Even-(2022-23)", "Odd-(2022-23)",
        "Supplementary 2021-22", "Re-ExamOdd (2021-22)", "Even-(2021-22)", "Odd-(2021-22)",
        "Supplementary 2020-21", "Even-(2020-21)", "Odd-(2020-21)",
        "Supplementary 2019-20", "Even-(2019-20)", "Odd-(2019-20)",
        "Special (2018-19)", "Even-(2018-19)", "Odd-(2018-19)",
        "Special-(2017-18)", "Even-(2017-18)", "Odd-(2017-18)",
        "Special-(2016-17)", "Even-(2016-17)", "Odd-(2016-17)",
        "Special-(2015-16)", "Even-(2015-16)", "Odd-(2015-16)"
    ]

    def __init__(self, timeout: int = 15):
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": f"{self.BASE_URL}/"
        })

    def get_available_sessions(self) -> List[str]:
        """Fetch available exam sessions dynamically from the BPUT results home page."""
        try:
            resp = self.session.get(f"{self.BASE_URL}/", timeout=self.timeout)
            if resp.status_code == 200:
                matches = re.findall(r'<option\s+value=[\'"]([^\'"]+)[\'"]', resp.text)
                sessions = [s.strip() for s in matches if s.strip() and s.strip() != "default_option"]
                if sessions:
                    return sessions
        except Exception:
            pass
        return list(self.FALLBACK_SESSIONS)

    def get_student_details(self, roll_no: str) -> Optional[Dict[str, Any]]:
        """
        Fetch basic student details (Name, College, Branch, Batch).
        Endpoint: POST /student-detsils-results?rollNo={rollNo}
        """
        url = f"{self.BASE_URL}/student-detsils-results?rollNo={roll_no.strip()}"
        try:
            resp = self.session.post(url, json={}, timeout=self.timeout)
            if resp.status_code == 200 and resp.text.strip():
                data = resp.json()
                if isinstance(data, dict) and data.get("rollNo"):
                    return data
        except Exception:
            pass
        return None

    def check_student_in_session(self, roll_no: str, session: str, dob: str = DEFAULT_DOB) -> List[Dict[str, Any]]:
        """
        Check if a student has published results in a given exam session.
        Endpoint: POST /student-results-list?rollNo={rollNo}&dob={dob}&session={session}
        Returns list of published semester entries for that session.
        """
        url = f"{self.BASE_URL}/student-results-list?rollNo={roll_no.strip()}&dob={dob.strip()}&session={session.strip()}"
        try:
            resp = self.session.post(url, json={}, timeout=self.timeout)
            if resp.status_code == 200 and resp.text.strip():
                data = resp.json()
                if isinstance(data, list):
                    return data
        except Exception:
            pass
        return []

    def get_semester_grades(self, roll_no: str, sem_id: Any, session: str) -> Dict[str, Any]:
        """
        Fetch full subject grades and SGPA for a specific semester and session.
        Endpoint: POST /student-results-subjects-list?semid={semId}&rollNo={rollNo}&session={session}
        Returns:
            {
                "grades": [...],
                "sgpadetails": {"cretits": int, "totalGradePoints": int, "sgpa": str}
            }
        """
        url = f"{self.BASE_URL}/student-results-subjects-list?semid={sem_id}&rollNo={roll_no.strip()}&session={session.strip()}"
        try:
            resp = self.session.post(url, json={}, timeout=self.timeout)
            if resp.status_code == 200 and resp.text.strip():
                data = resp.json()
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
        return {"grades": [], "sgpadetails": {}}

    def fetch_full_results(
        self,
        roll_no: str,
        dob: str = DEFAULT_DOB,
        sessions_to_check: Optional[List[str]] = None,
        max_workers: int = 10,
        progress_callback: Optional[Any] = None
    ) -> Dict[str, Any]:
        """
        Scrape complete academic records for a student across all sessions.

        Returns a dictionary containing:
            - 'student_info': Dict of student details
            - 'semesters': List of semester records (chronologically ordered)
            - 'summary': Cumulative metrics (CGPA, Total Credits, Total Grade Points, Backlogs)
        """
        roll_no = roll_no.strip()
        dob = dob.strip() if dob else self.DEFAULT_DOB

        # 1. Fetch Student Profile
        student_info = self.get_student_details(roll_no)

        # 2. Determine sessions to scan
        if not sessions_to_check:
            sessions_to_check = self.get_available_sessions()

        if progress_callback:
            progress_callback("sessions_start", len(sessions_to_check))

        # 3. Parallel scan across exam sessions to find published results
        session_results: List[Tuple[str, List[Dict[str, Any]]]] = []

        def _worker(sess: str):
            res = self.check_student_in_session(roll_no, sess, dob)
            if progress_callback:
                progress_callback("session_checked", sess)
            return (sess, res)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_worker, sess) for sess in sessions_to_check]
            for future in as_completed(futures):
                sess, res = future.result()
                if res:
                    session_results.append((sess, res))

        # 4. Fetch detailed subject grades for all found results
        # Flatten all semester appearances found
        sem_appearances: List[Dict[str, Any]] = []
        for sess, list_items in session_results:
            for item in list_items:
                sem_appearances.append({
                    "session": sess,
                    "semester": item.get("semester", ""),
                    "semId": item.get("semId", ""),
                    "branchName": item.get("branchName", ""),
                    "course": item.get("course", "")
                })

        if progress_callback:
            progress_callback("grades_start", len(sem_appearances))

        # Group appearances by semester to distinguish Regular vs Backlog/Improvement accurately
        sem_appearances_by_id: Dict[str, List[Dict[str, Any]]] = {}
        for app in sem_appearances:
            sid = str(app.get("semId", ""))
            sem_appearances_by_id.setdefault(sid, []).append(app)

        def _fetch_grades(appearance: Dict[str, Any]):
            details = self.get_semester_grades(roll_no, appearance["semId"], appearance["session"])
            if progress_callback:
                progress_callback("grades_fetched", appearance)
            return (appearance, details)

        detailed_semesters: List[Dict[str, Any]] = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            grade_futures = [executor.submit(_fetch_grades, app) for app in sem_appearances]
            for future in as_completed(grade_futures):
                app, details = future.result()
                grades = details.get("grades", [])
                sgpa_info = details.get("sgpadetails", {})

                sid = str(app.get("semId", ""))
                sess_lower = str(app["session"]).lower()
                all_for_this_sem = sem_appearances_by_id.get(sid, [])

                # Determine Exam Type:
                # 1. 8th semester B.Tech typically consists of a single Major Project / Internship (6-10 credits)
                if sid == "8" or str(app.get("semester", "")).lower().startswith("8"):
                    exam_type = "Regular (Project/Internship)"
                elif "supplementary" in sess_lower or "special" in sess_lower or "re-exam" in sess_lower:
                    exam_type = "Supplementary / Special"
                elif len(all_for_this_sem) > 1 and len(grades) <= 3:
                    exam_type = "Backlog / Improvement"
                else:
                    exam_type = "Regular"


                # Parse SGPA safely as float
                raw_sgpa = sgpa_info.get("sgpa")
                try:
                    sgpa_val = float(raw_sgpa) if raw_sgpa is not None else 0.0
                except (ValueError, TypeError):
                    sgpa_val = 0.0

                raw_credits = sgpa_info.get("cretits")
                try:
                    credits_val = int(raw_credits) if raw_credits is not None else 0
                except (ValueError, TypeError):
                    credits_val = 0

                raw_points = sgpa_info.get("totalGradePoints")
                try:
                    points_val = float(raw_points) if raw_points is not None else round(sgpa_val * credits_val, 2)
                except (ValueError, TypeError):
                    points_val = round(sgpa_val * credits_val, 2)

                # Check backlogs in this semester
                failed_subjects = [g for g in grades if str(g.get("grade", "")).strip().upper() in ["F", "M", "S"]]

                detailed_semesters.append({
                    "session": app["session"],
                    "semester": app["semester"],
                    "semId": app["semId"],
                    "course": app["course"],
                    "branchName": app["branchName"],
                    "exam_type": exam_type,
                    "grades": grades,
                    "credits": credits_val,
                    "points": points_val,
                    "sgpa": sgpa_val,
                    "sgpa_raw": raw_sgpa,
                    "failed_count": len(failed_subjects),
                    "status": "PASS" if len(failed_subjects) == 0 else f"{len(failed_subjects)} BACKLOG(S)"
                })

        # 5. Sort semesters logically:
        # Sort by semId (integer order: 1, 2, 3...) then by session
        def _sort_key(item):
            sem_id_num = 99
            try:
                sem_id_num = int(item["semId"])
            except Exception:
                # Try extracting digit from semester name (e.g., '1st' -> 1)
                m = re.search(r'\d+', str(item.get("semester", "")))
                if m:
                    sem_id_num = int(m.group())
            return (sem_id_num, item["session"])

        detailed_semesters.sort(key=_sort_key)

        # 6. Compute Overall Cumulative Statistics (CGPA, Total Credits)
        # BPUT CGPA Calculation:
        # For unique subjects, latest cleared grade points are considered.
        # Across distinct regular semesters: Sum(SGPA * Credits) / Sum(Credits)
        total_credits = 0
        total_grade_points = 0.0
        unique_subjects: Dict[str, Dict[str, Any]] = {}
        active_backlogs = set()

        for sem in detailed_semesters:
            for g in sem["grades"]:
                sub_code = str(g.get("subjectCODE", "")).strip().upper()
                grade = str(g.get("grade", "")).strip().upper()
                if sub_code:
                    unique_subjects[sub_code] = g
                    if grade in ["F", "M", "S"]:
                        active_backlogs.add(sub_code)
                    elif sub_code in active_backlogs:
                        # Cleared in a subsequent exam
                        active_backlogs.remove(sub_code)

        for sem in detailed_semesters:
            if sem["credits"] > 0:
                total_credits += sem["credits"]
                total_grade_points += sem["points"]

        overall_cgpa = round(total_grade_points / total_credits, 2) if total_credits > 0 else 0.0

        # If student_info was missing, try inferring from grades
        if not student_info and detailed_semesters:
            first_sem = detailed_semesters[0]
            student_info = {
                "rollNo": roll_no,
                "studentName": "N/A",
                "collegeName": "N/A",
                "branchName": first_sem.get("branchName", "N/A"),
                "courseName": first_sem.get("course", "N/A"),
                "batch": "N/A"
            }

        return {
            "roll_no": roll_no,
            "student_info": student_info or {
                "rollNo": roll_no,
                "studentName": "Unknown",
                "collegeName": "Unknown",
                "branchName": "Unknown",
                "courseName": "Unknown",
                "batch": "Unknown"
            },
            "semesters": detailed_semesters,
            "summary": {
                "total_semesters_found": len(detailed_semesters),
                "total_credits": total_credits,
                "total_grade_points": round(total_grade_points, 2),
                "cgpa": overall_cgpa,
                "total_unique_subjects": len(unique_subjects),
                "active_backlog_count": len(active_backlogs),
                "active_backlogs": list(active_backlogs),
                "academic_standing": "ALL CLEARED" if len(active_backlogs) == 0 else f"{len(active_backlogs)} PENDING BACKLOG(S)"
            }
        }
