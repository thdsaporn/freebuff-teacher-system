import sqlite3
import os
import json
from typing import List, Dict, Any, Optional

DB_PATH = os.path.join(os.path.dirname(__file__), "teachers.db")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Teachers table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS teachers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        prefix_rank TEXT DEFAULT '',
        first_name TEXT NOT NULL,
        last_name TEXT NOT NULL,
        age TEXT DEFAULT '',
        position TEXT DEFAULT '',
        affiliation TEXT DEFAULT '',
        workplace_address TEXT DEFAULT '',
        phone TEXT DEFAULT '',
        email TEXT DEFAULT '',
        photo_url TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Educations table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS educations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        year TEXT DEFAULT '',
        level TEXT DEFAULT '',
        degree_field TEXT DEFAULT '',
        institution TEXT DEFAULT '',
        FOREIGN KEY(teacher_id) REFERENCES teachers(id) ON DELETE CASCADE
    );
    """)

    # Trainings table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS trainings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        year TEXT DEFAULT '',
        course_name TEXT DEFAULT '',
        organized_by TEXT DEFAULT '',
        FOREIGN KEY(teacher_id) REFERENCES teachers(id) ON DELETE CASCADE
    );
    """)

    # Work histories table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS work_histories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        date_period TEXT DEFAULT '',
        position_role TEXT DEFAULT '',
        FOREIGN KEY(teacher_id) REFERENCES teachers(id) ON DELETE CASCADE
    );
    """)

    # Teaching assignments table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS teaching_assignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        course_type TEXT DEFAULT '',
        subject_category TEXT DEFAULT '',
        subject_name TEXT DEFAULT '',
        FOREIGN KEY(teacher_id) REFERENCES teachers(id) ON DELETE CASCADE
    );
    """)

    conn.commit()
    conn.close()

def get_all_teachers(query: Optional[str] = None, course_filter: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    sql = """
    SELECT DISTINCT t.* FROM teachers t
    LEFT JOIN teaching_assignments a ON t.id = a.teacher_id
    WHERE 1=1
    """
    params = []
    
    if query:
        q = f"%{query.strip()}%"
        sql += """ AND (
            t.prefix_rank LIKE ? OR
            t.first_name LIKE ? OR 
            t.last_name LIKE ? OR 
            t.position LIKE ? OR 
            t.affiliation LIKE ? OR
            a.subject_name LIKE ?
        )"""
        params.extend([q, q, q, q, q, q])
        
    if course_filter and course_filter != 'all':
        sql += " AND a.course_type LIKE ?"
        params.append(f"%{course_filter.strip()}%")
        
    sql += " ORDER BY t.id ASC"
    
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    
    teachers = []
    for r in rows:
        t = dict(r)
        t_id = t["id"]
        
        # Get courses summary
        cursor.execute("SELECT DISTINCT course_type FROM teaching_assignments WHERE teacher_id = ? AND course_type != ''", (t_id,))
        courses = [c["course_type"] for c in cursor.fetchall()]
        t["courses"] = courses
        
        # Get subjects summary
        cursor.execute("SELECT DISTINCT subject_name FROM teaching_assignments WHERE teacher_id = ? AND subject_name != ''", (t_id,))
        subjects = [s["subject_name"] for s in cursor.fetchall()]
        t["subjects"] = subjects
        
        # Get primary education
        cursor.execute("SELECT degree_field, institution FROM educations WHERE teacher_id = ? ORDER BY id DESC LIMIT 1", (t_id,))
        edu = cursor.fetchone()
        t["education_summary"] = f"{edu['degree_field']} {edu['institution']}".strip() if edu else ""
        
        teachers.append(t)
        
    conn.close()
    return teachers

def get_teacher_by_id(teacher_id: int) -> Optional[Dict[str, Any]]:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM teachers WHERE id = ?", (teacher_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None
        
    teacher = dict(row)
    
    # Educations
    cursor.execute("SELECT * FROM educations WHERE teacher_id = ? ORDER BY id ASC", (teacher_id,))
    teacher["educations"] = [dict(e) for e in cursor.fetchall()]
    
    # Trainings
    cursor.execute("SELECT * FROM trainings WHERE teacher_id = ? ORDER BY id ASC", (teacher_id,))
    teacher["trainings"] = [dict(tr) for tr in cursor.fetchall()]
    
    # Work histories
    cursor.execute("SELECT * FROM work_histories WHERE teacher_id = ? ORDER BY id ASC", (teacher_id,))
    teacher["work_histories"] = [dict(w) for w in cursor.fetchall()]
    
    # Teaching assignments
    cursor.execute("SELECT * FROM teaching_assignments WHERE teacher_id = ? ORDER BY id ASC", (teacher_id,))
    teacher["teaching_assignments"] = [dict(ta) for ta in cursor.fetchall()]
    
    conn.close()
    return teacher

def create_teacher(data: Dict[str, Any]) -> int:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    INSERT INTO teachers (
        prefix_rank, first_name, last_name, age, position, 
        affiliation, workplace_address, phone, email, photo_url, notes, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    """, (
        data.get("prefix_rank", "").strip(),
        data.get("first_name", "").strip(),
        data.get("last_name", "").strip(),
        data.get("age", "").strip(),
        data.get("position", "").strip(),
        data.get("affiliation", "").strip(),
        data.get("workplace_address", "").strip(),
        data.get("phone", "").strip(),
        data.get("email", "").strip(),
        data.get("photo_url", "").strip(),
        data.get("notes", "").strip()
    ))
    teacher_id = cursor.lastrowid
    
    # Insert educations
    for edu in data.get("educations", []):
        if any(str(v).strip() for v in edu.values()):
            cursor.execute("""
            INSERT INTO educations (teacher_id, year, level, degree_field, institution)
            VALUES (?, ?, ?, ?, ?)
            """, (teacher_id, edu.get("year", ""), edu.get("level", ""), edu.get("degree_field", ""), edu.get("institution", "")))
            
    # Insert trainings
    for tr in data.get("trainings", []):
        if any(str(v).strip() for v in tr.values()):
            cursor.execute("""
            INSERT INTO trainings (teacher_id, year, course_name, organized_by)
            VALUES (?, ?, ?, ?)
            """, (teacher_id, tr.get("year", ""), tr.get("course_name", ""), tr.get("organized_by", "")))
            
    # Insert work histories
    for w in data.get("work_histories", []):
        if any(str(v).strip() for v in w.values()):
            cursor.execute("""
            INSERT INTO work_histories (teacher_id, date_period, position_role)
            VALUES (?, ?, ?)
            """, (teacher_id, w.get("date_period", ""), w.get("position_role", "")))
            
    # Insert teaching assignments
    for ta in data.get("teaching_assignments", []):
        if any(str(v).strip() for v in ta.values()):
            cursor.execute("""
            INSERT INTO teaching_assignments (teacher_id, course_type, subject_category, subject_name)
            VALUES (?, ?, ?, ?)
            """, (teacher_id, ta.get("course_type", ""), ta.get("subject_category", ""), ta.get("subject_name", "")))
            
    conn.commit()
    conn.close()
    return teacher_id

def update_teacher(teacher_id: int, data: Dict[str, Any]) -> bool:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    UPDATE teachers SET
        prefix_rank = ?,
        first_name = ?,
        last_name = ?,
        age = ?,
        position = ?,
        affiliation = ?,
        workplace_address = ?,
        phone = ?,
        email = ?,
        photo_url = ?,
        notes = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    """, (
        data.get("prefix_rank", "").strip(),
        data.get("first_name", "").strip(),
        data.get("last_name", "").strip(),
        data.get("age", "").strip(),
        data.get("position", "").strip(),
        data.get("affiliation", "").strip(),
        data.get("workplace_address", "").strip(),
        data.get("phone", "").strip(),
        data.get("email", "").strip(),
        data.get("photo_url", "").strip(),
        data.get("notes", "").strip(),
        teacher_id
    ))
    
    # Replace educations
    cursor.execute("DELETE FROM educations WHERE teacher_id = ?", (teacher_id,))
    for edu in data.get("educations", []):
        if any(str(v).strip() for v in edu.values()):
            cursor.execute("""
            INSERT INTO educations (teacher_id, year, level, degree_field, institution)
            VALUES (?, ?, ?, ?, ?)
            """, (teacher_id, edu.get("year", ""), edu.get("level", ""), edu.get("degree_field", ""), edu.get("institution", "")))

    # Replace trainings
    cursor.execute("DELETE FROM trainings WHERE teacher_id = ?", (teacher_id,))
    for tr in data.get("trainings", []):
        if any(str(v).strip() for v in tr.values()):
            cursor.execute("""
            INSERT INTO trainings (teacher_id, year, course_name, organized_by)
            VALUES (?, ?, ?, ?)
            """, (teacher_id, tr.get("year", ""), tr.get("course_name", ""), tr.get("organized_by", "")))

    # Replace work histories
    cursor.execute("DELETE FROM work_histories WHERE teacher_id = ?", (teacher_id,))
    for w in data.get("work_histories", []):
        if any(str(v).strip() for v in w.values()):
            cursor.execute("""
            INSERT INTO work_histories (teacher_id, date_period, position_role)
            VALUES (?, ?, ?)
            """, (teacher_id, w.get("date_period", ""), w.get("position_role", "")))

    # Replace teaching assignments
    cursor.execute("DELETE FROM teaching_assignments WHERE teacher_id = ?", (teacher_id,))
    for ta in data.get("teaching_assignments", []):
        if any(str(v).strip() for v in ta.values()):
            cursor.execute("""
            INSERT INTO teaching_assignments (teacher_id, course_type, subject_category, subject_name)
            VALUES (?, ?, ?, ?)
            """, (teacher_id, ta.get("course_type", ""), ta.get("subject_category", ""), ta.get("subject_name", "")))

    conn.commit()
    conn.close()
    return True

def delete_teacher(teacher_id: int) -> bool:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM teachers WHERE id = ?", (teacher_id,))
    deleted = cursor.rowcount > 0
    conn.commit()
    conn.close()
    return deleted

def get_stats() -> Dict[str, Any]:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) as count FROM teachers")
    teacher_count = cursor.fetchone()["count"]
    
    cursor.execute("SELECT COUNT(DISTINCT course_type) as count FROM teaching_assignments WHERE course_type != ''")
    course_count = cursor.fetchone()["count"]
    
    cursor.execute("SELECT COUNT(DISTINCT subject_name) as count FROM teaching_assignments WHERE subject_name != ''")
    subject_count = cursor.fetchone()["count"]
    
    cursor.execute("SELECT DISTINCT course_type FROM teaching_assignments WHERE course_type != '' ORDER BY course_type")
    course_list = [r["course_type"] for r in cursor.fetchall()]
    
    conn.close()
    return {
        "teacher_count": teacher_count,
        "course_count": course_count,
        "subject_count": subject_count,
        "courses": course_list
    }
