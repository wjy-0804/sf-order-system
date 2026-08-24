-- ============================================================
-- 顺丰订单汇总系统 - Supabase / PostgreSQL 建表脚本
-- 用途：在 Supabase SQL Editor 中执行，初始化数据库结构
-- 说明：应用首次启动时 init_db() 也会自动建表（CREATE TABLE IF NOT EXISTS），
--       本脚本为手动预建/参考用，二者幂等，重复执行无副作用。
-- ============================================================

-- 用户表
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    is_admin INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
);

-- 订单表：status='pending' 未导出，'exported' 已导出
CREATE TABLE IF NOT EXISTS orders (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL,
    order_data TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    batch_id TEXT,
    created_at TEXT NOT NULL,
    exported_at TEXT,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- 地址簿分组表
CREATE TABLE IF NOT EXISTS address_groups (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- 地址簿明细表
CREATE TABLE IF NOT EXISTS address_book (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL,
    group_id INTEGER NOT NULL,
    name TEXT NOT NULL DEFAULT '',
    phone TEXT NOT NULL DEFAULT '',
    address TEXT NOT NULL DEFAULT '',
    product TEXT NOT NULL DEFAULT '',
    quantity TEXT NOT NULL DEFAULT '1',
    created_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id),
    FOREIGN KEY (group_id) REFERENCES address_groups(id) ON DELETE CASCADE
);

-- 导出存档表：每次导出时记录，保留72小时供重新下载
CREATE TABLE IF NOT EXISTS export_archives (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL,
    filename TEXT NOT NULL,
    file_path TEXT NOT NULL,
    order_count INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- 默认管理员账号：lmy123 / lmy123（超级管理员）
-- 应用启动时也会自动创建，这里作备份
INSERT INTO users (username, password, is_admin, created_at)
SELECT 'lmy123', 'lmy123', 1, to_char(now() AT TIME ZONE 'UTC', 'YYYY-MM-DD HH24:MI:SS')
WHERE NOT EXISTS (SELECT 1 FROM users WHERE username = 'lmy123');

-- 验证
SELECT 'users' AS table_name, count(*) AS rows FROM users
UNION ALL SELECT 'orders', count(*) FROM orders
UNION ALL SELECT 'address_groups', count(*) FROM address_groups
UNION ALL SELECT 'address_book', count(*) FROM address_book
UNION ALL SELECT 'export_archives', count(*) FROM export_archives;
