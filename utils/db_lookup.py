# utils/db_lookup.py

import os
from urllib.parse import urlparse

import pymysql
from dotenv import load_dotenv

load_dotenv()


def _resolve_db_config():
    """
    DB_URL / DB_USERNAME / DB_PASSWORD 우선 사용.
    기존 DB_HOST/DB_PORT/DB_USER/DB_NAME는 폴백으로 지원.
    """
    db_url = (os.environ.get("DB_URL") or "").strip()
    # Accept JDBC style URL as-is from env, e.g. jdbc:mysql://host:3306/db
    if db_url.lower().startswith("jdbc:"):
        db_url = db_url[5:]
    db_username = (os.environ.get("DB_USERNAME") or "").strip()
    db_password = os.environ.get("DB_PASSWORD", "")

    if db_url:
        parsed = urlparse(db_url)
        if parsed.scheme and parsed.hostname:
            return {
                "host": parsed.hostname,
                "port": parsed.port or 3306,
                "user": db_username or parsed.username or "root",
                "password": db_password or parsed.password or "",
                "db": (parsed.path or "/").lstrip("/") or os.environ.get("DB_NAME", "randi_db"),
            }

    return {
        "host": os.environ.get("DB_HOST", "127.0.0.1"),
        "port": int(os.environ.get("DB_PORT", 3306)),
        "user": os.environ.get("DB_USER", "root"),
        "password": db_password or os.environ.get("DB_PASSWORD", "rootpw"),
        "db": os.environ.get("DB_NAME", "randi_db"),
    }


def get_connection():
    """환경변수에서 DB 접속 정보를 읽어 MySQL 연결을 반환."""
    cfg = _resolve_db_config()
    return pymysql.connect(
        host=cfg["host"],
        port=int(cfg["port"]),
        user=cfg["user"],
        password=cfg["password"],
        db=cfg["db"],
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
    )


def get_notice_info_by_id(notice_id):
    """
    notice_id(PK)로 공고 정보 조회.
    Returns:
        dict: {"seq": "...", "author": "...", "title": "..."} or None
    """
    conn = None
    try:
        conn = get_connection()
        with conn.cursor() as cursor:
            sql = """
                SELECT
                    seq,
                    author,
                    title
                FROM project_notices
                WHERE notice_id = %s
                LIMIT 1
            """
            cursor.execute(sql, (notice_id,))
            row = cursor.fetchone()
            if row:
                return {
                    "seq": row["seq"],
                    "author": row["author"],
                    "title": row.get("title", ""),
                }
    except Exception as e:
        print(f"[DB Error] get_notice_info_by_id 실패: {e}")
        return None
    finally:
        if conn:
            conn.close()
    return None


def find_ministry_by_seq_author(seq, author=None):
    """
    공고 번호(seq)로 부처명을 찾는다.
    author가 이미 있으면 그대로 반환.
    """
    if author:
        return author

    conn = None
    try:
        conn = get_connection()
        with conn.cursor() as cursor:
            sql = "SELECT author FROM project_notices WHERE seq = %s LIMIT 1"
            cursor.execute(sql, (seq,))
            result = cursor.fetchone()
            if result:
                return result["author"]
    except Exception as e:
        print(f"[DB 조회 실패] find_ministry_by_seq_author: {e}")
        return None
    finally:
        if conn:
            conn.close()
    return None
