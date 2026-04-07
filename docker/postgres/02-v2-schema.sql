-- V2 Source Tables

CREATE TABLE IF NOT EXISTS workbook (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    owner_user_id TEXT NOT NULL,
    head_revision INTEGER NOT NULL DEFAULT 0,
    calculated_revision INTEGER NOT NULL DEFAULT 0,
    calc_mode TEXT NOT NULL DEFAULT 'automatic',
    compatibility_mode TEXT NOT NULL DEFAULT 'excel-modern',
    recalc_epoch INTEGER NOT NULL DEFAULT 0,
    created_at BIGINT NOT NULL,
    updated_at BIGINT NOT NULL
);

CREATE TABLE IF NOT EXISTS workbook_member (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL,
    role TEXT NOT NULL,
    joined_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, user_id)
);

CREATE TABLE IF NOT EXISTS sheet (
    id TEXT NOT NULL,
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    position INTEGER NOT NULL,
    freeze_rows INTEGER NOT NULL DEFAULT 0,
    freeze_cols INTEGER NOT NULL DEFAULT 0,
    tab_color TEXT,
    hidden BOOLEAN DEFAULT FALSE,
    created_at BIGINT NOT NULL,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS cell_style (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    style_id TEXT NOT NULL,
    style_json JSONB NOT NULL,
    hash TEXT NOT NULL,
    created_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, style_id)
);

CREATE TABLE IF NOT EXISTS number_format (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    format_id TEXT NOT NULL,
    kind TEXT NOT NULL,
    code TEXT NOT NULL,
    created_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, format_id)
);

CREATE TABLE IF NOT EXISTS cell_input (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    sheet_id TEXT NOT NULL REFERENCES sheet(id) ON DELETE CASCADE,
    row_num INTEGER NOT NULL,
    col_num INTEGER NOT NULL,
    address TEXT NOT NULL,
    input_json JSONB,
    formula_source TEXT,
    style_id TEXT,
    format_id TEXT,
    editor_text TEXT,
    source_revision INTEGER NOT NULL,
    updated_by TEXT NOT NULL,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, sheet_id, address)
);

CREATE TABLE IF NOT EXISTS sheet_row (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    sheet_id TEXT NOT NULL REFERENCES sheet(id) ON DELETE CASCADE,
    row_num INTEGER NOT NULL,
    axis_id TEXT,
    height INTEGER,
    hidden BOOLEAN NOT NULL DEFAULT FALSE,
    outline_level INTEGER NOT NULL DEFAULT 0,
    source_revision INTEGER NOT NULL,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, sheet_id, row_num)
);

CREATE TABLE IF NOT EXISTS sheet_col (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    sheet_id TEXT NOT NULL REFERENCES sheet(id) ON DELETE CASCADE,
    col_num INTEGER NOT NULL,
    axis_id TEXT,
    width INTEGER,
    hidden BOOLEAN NOT NULL DEFAULT FALSE,
    outline_level INTEGER NOT NULL DEFAULT 0,
    source_revision INTEGER NOT NULL,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, sheet_id, col_num)
);

CREATE TABLE IF NOT EXISTS defined_name (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    scope_sheet_id TEXT,
    name TEXT NOT NULL,
    normalized_name TEXT NOT NULL,
    value_json JSONB NOT NULL,
    source_revision INTEGER NOT NULL,
    PRIMARY KEY (workbook_id, scope_sheet_id, normalized_name)
);

CREATE TABLE IF NOT EXISTS workbook_change (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    change_id TEXT NOT NULL,
    revision INTEGER NOT NULL,
    actor_user_id TEXT NOT NULL,
    kind TEXT NOT NULL,
    sheet_id TEXT,
    range_json JSONB,
    summary_json JSONB NOT NULL,
    before_json JSONB,
    after_json JSONB,
    created_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, change_id)
);

-- V2 Render Tables

CREATE TABLE IF NOT EXISTS cell_render (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    sheet_id TEXT NOT NULL REFERENCES sheet(id) ON DELETE CASCADE,
    row_num INTEGER NOT NULL,
    col_num INTEGER NOT NULL,
    address TEXT NOT NULL,
    value_tag TEXT NOT NULL,
    number_value DOUBLE PRECISION,
    string_value TEXT,
    boolean_value BOOLEAN,
    error_code TEXT,
    style_id TEXT,
    format_id TEXT,
    flags INTEGER NOT NULL DEFAULT 0,
    calc_revision INTEGER NOT NULL,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, sheet_id, address)
);

CREATE TABLE IF NOT EXISTS presence (
    workbook_id TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    session_id TEXT NOT NULL,
    user_id TEXT NOT NULL,
    sheet_id TEXT,
    address TEXT,
    selection_json JSONB,
    updated_at BIGINT NOT NULL,
    PRIMARY KEY (workbook_id, session_id)
);

-- Publication for Zero v2
DROP PUBLICATION IF EXISTS zero_data_v2;
CREATE PUBLICATION zero_data_v2 FOR TABLE 
    workbook, 
    workbook_member, 
    sheet, 
    cell_style, 
    number_format, 
    cell_input, 
    sheet_row, 
    sheet_col, 
    defined_name, 
    workbook_change, 
    cell_render,
    presence;
