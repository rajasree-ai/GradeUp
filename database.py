import os
import sqlite3
from datetime import datetime

DB_PATH = os.getenv("SQLITE_DB_PATH", "results.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()


def get_conn():
    if DATABASE_URL:
        import psycopg
        from psycopg.rows import dict_row

        conninfo = DATABASE_URL
        if conninfo.startswith("postgres://"):
            conninfo = conninfo.replace("postgres://", "postgresql://", 1)
        return psycopg.connect(
            conninfo,
            row_factory=dict_row,
            connect_timeout=10,
        )

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _uses_postgres():
    return bool(DATABASE_URL)


def _sql(sql):
    return sql.replace("?", "%s") if _uses_postgres() else sql


def init_db():
    conn = get_conn()
    c = conn.cursor()

    if _uses_postgres():
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id BIGSERIAL PRIMARY KEY,
                full_name TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                department TEXT,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS uploads (
                id BIGSERIAL PRIMARY KEY,
                user_id BIGINT REFERENCES users(id) ON DELETE SET NULL,
                filename TEXT NOT NULL,
                pass_mark INTEGER NOT NULL,
                uploaded_at TEXT NOT NULL,
                total_students INTEGER,
                passed INTEGER,
                failed INTEGER
            )
        """)
    else:
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                department TEXT,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS uploads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                filename TEXT NOT NULL,
                pass_mark INTEGER NOT NULL,
                uploaded_at TEXT NOT NULL,
                total_students INTEGER,
                passed INTEGER,
                failed INTEGER,
                FOREIGN KEY (user_id) REFERENCES users(id)
            )
        """)
    conn.commit()
    conn.close()


def create_user(full_name, email, department, password_hash):
    conn = get_conn()
    c = conn.cursor()
    try:
        query = """
            INSERT INTO users (full_name, email, department, password_hash, created_at)
            VALUES (?, ?, ?, ?, ?)
        """
        if _uses_postgres():
            query += " RETURNING id"
        c.execute(_sql(query), (full_name, email.lower().strip(), department or "",
                                password_hash, datetime.now().isoformat()))
        conn.commit()
        if _uses_postgres():
            return c.fetchone()["id"]
        return c.lastrowid
    except sqlite3.IntegrityError:
        conn.rollback()
        return None
    except Exception as exc:
        try:
            import psycopg
            if isinstance(exc, (psycopg.errors.UniqueViolation,
                                psycopg.errors.IntegrityError)):
                conn.rollback()
                return None
        except ImportError:
            pass
        conn.rollback()
        raise
    finally:
        conn.close()


def get_user_by_email(email):
    conn = get_conn()
    c = conn.cursor()
    c.execute(_sql("SELECT * FROM users WHERE email = ?"),
              (email.lower().strip(),))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_id(user_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute(_sql("SELECT * FROM users WHERE id = ?"), (user_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def save_upload(user_id, filename, pass_mark, stats):
    conn = get_conn()
    c = conn.cursor()
    c.execute(_sql("""
        INSERT INTO uploads (user_id, filename, pass_mark, uploaded_at,
                             total_students, passed, failed)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """), (user_id, filename, pass_mark, datetime.now().isoformat(),
           stats["total_students"], stats["passed_students"], stats["failed_students"]))
    conn.commit()
    conn.close()


def get_history(user_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute(_sql("""
        SELECT * FROM uploads WHERE user_id = ?
        ORDER BY id DESC LIMIT 20
    """), (user_id,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows