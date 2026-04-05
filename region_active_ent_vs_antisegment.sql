-- Requires:
--   /Users/grigorijkrasovickij/Documents/Playground/region_active_products.sql
--   and created table usr_wrk.region_active_products_long
--
-- Output table:
--   usr_wrk.region_active_ent_vs_anti_long

drop table if exists usr_wrk.region_active_ent_vs_anti_long;
create table usr_wrk.region_active_ent_vs_anti_long as
with ent as (
    select
        dt,
        month_end,
        bucket,
        russian_region_nm,
        metric_type,
        metric_nm,
        metric_value as ent_metric_value,
        active_cnt as ent_active_cnt,
        penetration as ent_penetration,
        total_metric_value as ent_total_metric_value,
        total_active_cnt as ent_total_active_cnt,
        portfolio_penetration as ent_portfolio_penetration
    from usr_wrk.region_active_products_long
    where segment = 'Активные предприниматели'
),
anti_region as (
    select
        dt,
        month_end,
        bucket,
        russian_region_nm,
        metric_type,
        metric_nm,
        metric_value as anti_metric_value_region,
        active_cnt as anti_active_cnt_region,
        penetration as anti_penetration_region
    from usr_wrk.region_active_products_long
    where segment = 'Активный антисегмент'
),
anti_portfolio as (
    select distinct
        dt,
        month_end,
        metric_type,
        metric_nm,
        total_metric_value as anti_total_metric_value,
        total_active_cnt as anti_total_active_cnt,
        portfolio_penetration as anti_portfolio_penetration
    from usr_wrk.region_active_products_long
    where segment = 'Активный антисегмент'
)
select
    e.dt,
    e.month_end,
    e.bucket,
    e.russian_region_nm,
    e.metric_type,
    e.metric_nm,
    e.ent_metric_value,
    e.ent_active_cnt,
    e.ent_penetration,
    e.ent_total_metric_value,
    e.ent_total_active_cnt,
    e.ent_portfolio_penetration,
    a.anti_metric_value_region,
    a.anti_active_cnt_region,
    a.anti_penetration_region,
    ap.anti_total_metric_value,
    ap.anti_total_active_cnt,
    ap.anti_portfolio_penetration,
    round((e.ent_penetration / nullif(a.anti_penetration_region, 0) - 1), 4) as deviation_vs_antisegment_region,
    round(((e.ent_penetration / nullif(a.anti_penetration_region, 0) - 1) * 100), 2) as deviation_vs_antisegment_region_pct,
    round((e.ent_penetration / nullif(ap.anti_portfolio_penetration, 0) - 1), 4) as deviation_vs_antisegment_portfolio,
    round(((e.ent_penetration / nullif(ap.anti_portfolio_penetration, 0) - 1) * 100), 2) as deviation_vs_antisegment_portfolio_pct
from ent e
left join anti_region a
    on e.dt = a.dt
   and e.month_end = a.month_end
   and e.metric_type = a.metric_type
   and e.metric_nm = a.metric_nm
   and e.bucket is not distinct from a.bucket
   and e.russian_region_nm is not distinct from a.russian_region_nm
left join anti_portfolio ap
    on e.dt = ap.dt
   and e.month_end = ap.month_end
   and e.metric_type = ap.metric_type
   and e.metric_nm = ap.metric_nm
distributed by (dt, metric_nm);
