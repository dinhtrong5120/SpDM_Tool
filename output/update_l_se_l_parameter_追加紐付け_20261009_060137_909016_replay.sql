-- Snapshot replay generated from:
--   update_l_se_l_parameter_追加紐付け_20261009_060137_909016.xlsx
-- Source tool: utils/Tools_new/update_l_se.py
-- Report status: UPDATED
-- Report database/host: spdm / localhost
-- Report result: 4 l_parameter rows and 19 prj_rfl rows committed.
--
-- This script replays only the exact UPDATE rows recorded by the report.
-- Old-state guards use MySQL's NULL-safe equality operator (<=>).
-- When run against the database already updated by the reported execution,
-- the guarded statements are expected to affect zero rows.
--
-- Reported non-mutations intentionally omitted:
--   * prj_rfl (1010, 3, 13808): SKIP_NO_VALUE.
--   * l_parameter 1000384: no linked RFL row (SKIP_NO_RFL); its own
--     related_se_parameter_id update is still included below.

SET NAMES utf8mb4;
USE `spdm`;

SELECT DATABASE() AS connected_database, 'spdm' AS expected_database;

START TRANSACTION;

-- -------------------------------------------------------------------------
-- l_parameter: expected report mutations = 4
-- -------------------------------------------------------------------------

UPDATE `l_parameter`
SET `related_se_parameter_id` = 262
WHERE `id` = 1000095
  AND `related_se_parameter_id` <=> NULL;
SET @rc_l_1000095 := ROW_COUNT();

UPDATE `l_parameter`
SET `related_se_parameter_id` = 635
WHERE `id` = 1000384
  AND `related_se_parameter_id` <=> NULL;
SET @rc_l_1000384 := ROW_COUNT();

UPDATE `l_parameter`
SET `related_se_parameter_id` = 877
WHERE `id` = 1000440
  AND `related_se_parameter_id` <=> NULL;
SET @rc_l_1000440 := ROW_COUNT();

UPDATE `l_parameter`
SET `related_se_parameter_id` = 711
WHERE `id` = 1001133
  AND `related_se_parameter_id` <=> 216;
SET @rc_l_1001133 := ROW_COUNT();

-- -------------------------------------------------------------------------
-- prj_rfl: expected report mutations = 19
-- Non-ASCII values use UTF-8 hexadecimal literals to avoid SQL-mode escaping
-- ambiguity. All reported map_name_before values below are NULL.
-- -------------------------------------------------------------------------

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 30 AND `rfl_id` = 16592
  AND `map_name` <=> NULL;
SET @rc_r_1000_30_16592 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 30 AND `rfl_id` = 20031
  AND `map_name` <=> NULL;
SET @rc_r_1000_30_20031 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 30 AND `rfl_id` = 22592
  AND `map_name` <=> NULL;
SET @rc_r_1000_30_22592 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 36 AND `rfl_id` = 20031
  AND `map_name` <=> NULL;
SET @rc_r_1000_36_20031 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 36 AND `rfl_id` = 32101
  AND `map_name` <=> NULL;
SET @rc_r_1000_36_32101 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1000 AND `phase_id` = 36 AND `rfl_id` = 1048928
  AND `map_name` <=> NULL;
SET @rc_r_1000_36_1048928 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = 'N/A'
WHERE `project_info_id` = 1001 AND `phase_id` = 28 AND `rfl_id` = 14515
  AND `map_name` <=> NULL;
SET @rc_r_1001_28_14515 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = 'N/A'
WHERE `project_info_id` = 1001 AND `phase_id` = 28 AND `rfl_id` = 14932
  AND `map_name` <=> NULL;
SET @rc_r_1001_28_14932 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = 'N/A'
WHERE `project_info_id` = 1001 AND `phase_id` = 28 AND `rfl_id` = 14993
  AND `map_name` <=> NULL;
SET @rc_r_1001_28_14993 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = 'N/A'
WHERE `project_info_id` = 1011 AND `phase_id` = 6 AND `rfl_id` = 22701
  AND `map_name` <=> NULL;
SET @rc_r_1011_6_22701 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1014 AND `phase_id` = 31 AND `rfl_id` = 20031
  AND `map_name` <=> NULL;
SET @rc_r_1014_31_20031 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1014 AND `phase_id` = 31 AND `rfl_id` = 32101
  AND `map_name` <=> NULL;
SET @rc_r_1014_31_32101 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1014 AND `phase_id` = 31 AND `rfl_id` = 1048928
  AND `map_name` <=> NULL;
SET @rc_r_1014_31_1048928 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1015 AND `phase_id` = 30 AND `rfl_id` = 20031
  AND `map_name` <=> NULL;
SET @rc_r_1015_30_20031 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1015 AND `phase_id` = 30 AND `rfl_id` = 31899
  AND `map_name` <=> NULL;
SET @rc_r_1015_30_31899 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
WHERE `project_info_id` = 1015 AND `phase_id` = 30 AND `rfl_id` = 31988
  AND `map_name` <=> NULL;
SET @rc_r_1015_30_31988 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1016 AND `phase_id` = 32 AND `rfl_id` = 20031
  AND `map_name` <=> NULL;
SET @rc_r_1016_32_20031 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1016 AND `phase_id` = 32 AND `rfl_id` = 22592
  AND `map_name` <=> NULL;
SET @rc_r_1016_32_22592 := ROW_COUNT();

UPDATE `prj_rfl`
SET `map_name` = CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
WHERE `project_info_id` = 1016 AND `phase_id` = 32 AND `rfl_id` = 31899
  AND `map_name` <=> NULL;
SET @rc_r_1016_32_31899 := ROW_COUNT();

-- Mutation counts. Each row has expected_rows = 1 when replaying the old
-- snapshot. A changed_rows value of 0 is expected if the UPDATED report has
-- already been applied to this database.
SELECT mutation, expected_rows, changed_rows
FROM (
    SELECT 'l_parameter:1000095' AS mutation, 1 AS expected_rows, @rc_l_1000095 AS changed_rows
    UNION ALL SELECT 'l_parameter:1000384', 1, @rc_l_1000384
    UNION ALL SELECT 'l_parameter:1000440', 1, @rc_l_1000440
    UNION ALL SELECT 'l_parameter:1001133', 1, @rc_l_1001133
    UNION ALL SELECT 'prj_rfl:1000:30:16592', 1, @rc_r_1000_30_16592
    UNION ALL SELECT 'prj_rfl:1000:30:20031', 1, @rc_r_1000_30_20031
    UNION ALL SELECT 'prj_rfl:1000:30:22592', 1, @rc_r_1000_30_22592
    UNION ALL SELECT 'prj_rfl:1000:36:20031', 1, @rc_r_1000_36_20031
    UNION ALL SELECT 'prj_rfl:1000:36:32101', 1, @rc_r_1000_36_32101
    UNION ALL SELECT 'prj_rfl:1000:36:1048928', 1, @rc_r_1000_36_1048928
    UNION ALL SELECT 'prj_rfl:1001:28:14515', 1, @rc_r_1001_28_14515
    UNION ALL SELECT 'prj_rfl:1001:28:14932', 1, @rc_r_1001_28_14932
    UNION ALL SELECT 'prj_rfl:1001:28:14993', 1, @rc_r_1001_28_14993
    UNION ALL SELECT 'prj_rfl:1011:6:22701', 1, @rc_r_1011_6_22701
    UNION ALL SELECT 'prj_rfl:1014:31:20031', 1, @rc_r_1014_31_20031
    UNION ALL SELECT 'prj_rfl:1014:31:32101', 1, @rc_r_1014_31_32101
    UNION ALL SELECT 'prj_rfl:1014:31:1048928', 1, @rc_r_1014_31_1048928
    UNION ALL SELECT 'prj_rfl:1015:30:20031', 1, @rc_r_1015_30_20031
    UNION ALL SELECT 'prj_rfl:1015:30:31899', 1, @rc_r_1015_30_31899
    UNION ALL SELECT 'prj_rfl:1015:30:31988', 1, @rc_r_1015_30_31988
    UNION ALL SELECT 'prj_rfl:1016:32:20031', 1, @rc_r_1016_32_20031
    UNION ALL SELECT 'prj_rfl:1016:32:22592', 1, @rc_r_1016_32_22592
    UNION ALL SELECT 'prj_rfl:1016:32:31899', 1, @rc_r_1016_32_31899
) AS mutation_counts;

SELECT
    'l_parameter' AS table_name,
    'UPDATE' AS action,
    4 AS expected_rows,
    (@rc_l_1000095 + @rc_l_1000384 + @rc_l_1000440 + @rc_l_1001133) AS changed_rows
UNION ALL
SELECT
    'prj_rfl',
    'UPDATE',
    19,
    (@rc_r_1000_30_16592 + @rc_r_1000_30_20031 + @rc_r_1000_30_22592
     + @rc_r_1000_36_20031 + @rc_r_1000_36_32101 + @rc_r_1000_36_1048928
     + @rc_r_1001_28_14515 + @rc_r_1001_28_14932 + @rc_r_1001_28_14993
     + @rc_r_1011_6_22701
     + @rc_r_1014_31_20031 + @rc_r_1014_31_32101 + @rc_r_1014_31_1048928
     + @rc_r_1015_30_20031 + @rc_r_1015_30_31899 + @rc_r_1015_30_31988
     + @rc_r_1016_32_20031 + @rc_r_1016_32_22592 + @rc_r_1016_32_31899);

-- Post-change verification: all four l_parameter rows must be present at the
-- exact values recorded as the report's after-state.
SELECT
    expected.id,
    expected.related_se_parameter_id AS expected_related_se_parameter_id,
    actual.related_se_parameter_id AS actual_related_se_parameter_id,
    CASE
        WHEN actual.id IS NULL THEN 'MISSING'
        WHEN actual.related_se_parameter_id <=> expected.related_se_parameter_id THEN 'OK'
        ELSE 'MISMATCH'
    END AS verification
FROM (
    SELECT 1000095 AS id, 262 AS related_se_parameter_id
    UNION ALL SELECT 1000384, 635
    UNION ALL SELECT 1000440, 877
    UNION ALL SELECT 1001133, 711
) AS expected
LEFT JOIN `l_parameter` AS actual ON actual.id = expected.id
ORDER BY expected.id;

-- Post-change verification: all 19 prj_rfl composite keys must be present at
-- the exact map_name values recorded as the report's after-state.
SELECT
    expected.project_info_id,
    expected.phase_id,
    expected.rfl_id,
    expected.map_name AS expected_map_name,
    actual.map_name AS actual_map_name,
    CASE
        WHEN actual.rfl_id IS NULL THEN 'MISSING'
        WHEN actual.map_name <=> expected.map_name THEN 'OK'
        ELSE 'MISMATCH'
    END AS verification
FROM (
    SELECT 1000 AS project_info_id, 30 AS phase_id, 16592 AS rfl_id,
           CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4) AS map_name
    UNION ALL SELECT 1000, 30, 20031, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
    UNION ALL SELECT 1000, 30, 22592, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
    UNION ALL SELECT 1000, 36, 20031, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1000, 36, 32101, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1000, 36, 1048928, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1001, 28, 14515, 'N/A'
    UNION ALL SELECT 1001, 28, 14932, 'N/A'
    UNION ALL SELECT 1001, 28, 14993, 'N/A'
    UNION ALL SELECT 1011, 6, 22701, 'N/A'
    UNION ALL SELECT 1014, 31, 20031, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1014, 31, 32101, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1014, 31, 1048928, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1015, 30, 20031, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
    UNION ALL SELECT 1015, 30, 31899, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
    UNION ALL SELECT 1015, 30, 31988, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303031 USING utf8mb4)
    UNION ALL SELECT 1016, 32, 20031, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1016, 32, 22592, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
    UNION ALL SELECT 1016, 32, 31899, CONVERT(0x3236303632325fe382a8e3838de3839ee3838de5a489e9809fe7b79a5f53715f4672616d655f44796e616d69635f525f302e3430325b6d5d5f46696e616c5f476561725f526174696f5f332e3335375f303032 USING utf8mb4)
) AS expected
LEFT JOIN `prj_rfl` AS actual
  ON actual.project_info_id = expected.project_info_id
 AND actual.phase_id = expected.phase_id
 AND actual.rfl_id = expected.rfl_id
ORDER BY expected.project_info_id, expected.phase_id, expected.rfl_id;

-- Review the mutation counts and verification results above, then run exactly
-- one of the following final choices on this same connection:
-- COMMIT;
-- ROLLBACK;
