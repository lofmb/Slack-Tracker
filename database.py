import sqlite3
from datetime import datetime

DB_NAME = "trackbot.db"

def get_connection():
    # Open a SQLite database connection and return the connection object.
    # Uses the global DB_NAME value and configures the connection so rows
    # can be accessed like dictionaries via sqlite3.Row.

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def setup_database():
    # Create the database schema if it does not already exist.
    # This function creates two tables:
    #   - tasks: stores task metadata, status, and elapsed time
    #   - time_segments: stores individual work segments for a task
    # It commits the schema changes and closes the database connection.
    # From Phase 1 Field Sheeting is from field_design to field_elapsed
    # Phase 2 Border Sheeting is border_design to border_elapsed
    # Phase 3 Packing is packing_elapsed
    # Phase 4 Notes is general_notes to issues_encountered

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            task_id     INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id     TEXT NOT NULL,
            channel_id  TEXT NOT NULL,
            customer_name TEXT NOT NULL,
            invoice_number TEXT NOT NULL,
            task_description TEXT NOT NULL,
            due_date     TEXT,
            is_na_due_date INTEGER DEFAULT 0,
            field_design TEXT,
            difficulty   TEXT,
            field_elapsed INTEGER DEFAULT 0,
            border_design   TEXT,
            border_difficulty TEXT,
            border_elapsed INTEGER DEFAULT 0,
            packing_elapsed INTEGER DEFAULT 0,
            general_notes TEXT,
            issues_encountered  TEXT,
            status       TEXT NOT NULL DEFAULT 'created',
            created_at   TEXT NOT NULL,
            message_ts  TEXT,
            total_elapsed INTEGER DEFAULT 0
        )
        """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS time_segments(
            segment_id      INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id         INTEGER NOT NULL,
            phase           TEXT NOT NULL,
            started_at      TEXT NOT NULL,
            stopped_at      TEXT,
            FOREIGN KEY (task_id) REFERENCES tasks (task_id)
        )
        """)

    conn.commit()
    conn.close()
    print("Database Ready.")
    
def create_task(user_id, channel_id, customer_name, invoice_number, task_description, due_date, is_na, design, difficulty):
    # Insert a new task record and return its generated task_id.
    # The task is created with the current timestamp and an initial status
    # of 'open' unless the table default overrides it.

    conn = get_connection()
    cursor = conn.cursor()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT INTO tasks (user_id, channel_id, customer_name, invoice_number, task_description, due_date, is_na_due_date, field_design, difficulty, status, current_phase, created_at)
        VALUES (?,?,?,?,?,?,?,?,?,'created', 'field_sheeting', ?)
        """, (user_id, channel_id, customer_name, invoice_number, task_description, due_date, is_na, design, difficulty, created_at)
        )

    conn.commit()
    task_id = cursor.lastrowid
    conn.close()
    return task_id

def get_active_task(user_id):
    # Return the most recent active task for a user, if one exists.
    # An active task is any task with status 'created', 'in_progress', or 'paused'.

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM tasks
        WHERE user_id = ? AND status IN ('created', 'in_progress', 'paused')
        ORDER BY task_id DESC
        LIMIT 1 
        """, (user_id,)
        )

    task = cursor.fetchone()
    conn.close()
    return task
    
def get_task (task_id):
    # Retrieve a single task by its task_id.

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM tasks WHERE task_id = ?", (task_id,))
    task = cursor.fetchone()
    conn.close()
    return task

def start_task (task_id):
    # Mark a task as started and create a new time segment.
    # This sets the task's status to 'in_progress' and records the time the
    # segment began.

    conn = get_connection()
    cursor = conn.cursor()
    started_at =  datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cursor.execute("SELECT current_phase FROM tasks WHERE task_id = ?",(task_id,))
    row = cursor.fetchone()
    current_phase = row["current_phase"]

    cursor.execute("""
        UPDATE tasks SET status = 'in_progress' WHERE task_id = ?
        """, (task_id,))

    cursor.execute("""
        INSERT INTO time_segments (task_id, phase, started_at) VALUES (?,?,?)
        """, (task_id, current_phase, started_at))

    conn.commit()
    conn.close()
    
def stop_task(task_id):
    # Stop the current active time segment and pause the task.
    # This updates the latest open time segment's stopped_at timestamp,
    # recalculates the total elapsed seconds for the task, and sets the
    # task status to 'paused'.

    conn = get_connection()
    cursor = conn.cursor()
    stopped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        UPDATE time_segments SET stopped_at = ?
        WHERE task_id = ? AND stopped_at IS NULL
        """, (stopped_at, task_id))

    cursor.execute("""
        SELECT started_at, stopped_at FROM time_segments
        WHERE task_id = ? AND stopped_at IS NOT NULL
        """, (task_id,))

    segments = cursor.fetchall()
    total = 0
    for seg in segments:
        start = datetime.strptime(seg["started_at"], "%Y-%m-%d %H:%M:%S")
        stop = datetime.strptime(seg["stopped_at"], "%Y-%m-%d %H:%M:%S")
        total += int((stop - start).total_seconds())
        
    cursor.execute("""
        UPDATE tasks SET status = 'paused', total_elapsed = ? WHERE task_id = ? """,
        (total, task_id))

    conn.commit()
    conn.close()
    
def complete_task(task_id):
    # Complete a task by stopping any open time segment and marking it done.
    # The task status becomes 'completed' and the task's total_elapsed field is
    # updated with the sum of all recorded segment durations.

    conn = get_connection()
    cursor = conn.cursor()
    stopped_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        UPDATE time_segments SET stopped_at = ?
        WHERE task_id = ? AND stopped_at IS NULL
        """,
        (stopped_at, task_id))

    cursor.execute("""
        SELECT started_at, stopped_at FROM time_segments
        WHERE task_id = ? AND stopped_at IS NOT NULL
        """, (task_id,))

    segment = cursor.fetchall()
    total = 0
    for seg in segment:
        start = datetime.strptime(seg["started_at"], "%Y-%m-%d %H:%M:%S")
        stop = datetime.strptime(seg["stopped_at"], "%Y-%m-%d %H:%M:%S")
        total += int((stop - start).total_seconds())

    cursor.execute("""
        UPDATE tasks SET status = 'completed', total_elapsed = ? WHERE task_id = ? """, (total, task_id))

    conn.commit()
    conn.close()
    
def update_message_ts(task_id, message_ts):
    # Store the Slack message timestamp for a task.
    # This updates the tasks table field used to track the last message sent for
    # a task update or notification.

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE tasks SET message_ts = ? WHERE task_id = ? """, (message_ts, task_id))

    conn.commit()
    conn.close()
    
def format_elapsed(seconds):
    # Format elapsed seconds into a readable hours/minutes/seconds string.

    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours} h {minutes}m {secs}s"

def delete_task(task_id):
    conn = get_connection()
    cursor = conn.cursor()
    # remove any associated time segments first
    cursor.execute("DELETE FROM time_segments WHERE task_id = ?", (task_id,))
    # remove the task by its primary key `task_id`
    cursor.execute("DELETE FROM tasks WHERE task_id = ?", (task_id,))
    conn.commit()
    conn.close()
    
def update_task(task_id, customer, invoice, task_desc, design, difficulty, due_date):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tasks 
        SET customer_name = ?, 
            invoice_number = ?, 
            task_description = ?, 
            field_design = ?, 
            difficulty = ?, 
            due_date = ?
        WHERE task_id = ?
    """, (customer, invoice, task_desc, design, difficulty, due_date, task_id))
    conn.commit()
    conn.close()

if __name__ == "__main__":
    setup_database()