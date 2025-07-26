CREATE DEFINER=`root`@`%` PROCEDURE `get__refunded_orders_data`(IN p_order_id BIGINT)
BEGIN
    SELECT
       opl.order_id,
       opl.is_refund,
       opl.parent_order_id,
       opl.product_id,
       wpp.post_title,
       opl.price,
       opl.is_bundle,
       opl.order_item_id,
       opl.item_id,
       opl.parent_product_id,
       opl.parent_item_id,
       wppp.post_title AS parent_product_title,
       opl.start_date,
       opl.end_date,
       opl.used_days,
       opl.discount_days,
       opl.product_qty,
       aop.payment_amount,
       aop.payment_type,
       aop.date,
       aop.description
    FROM
        wp_wc_order_product_lookup opl
    LEFT JOIN wp_posts wpp ON wpp.id=opl.product_id
    LEFT JOIN wp_posts wppp ON wppp.id=opl.parent_product_id
    LEFT JOIN app_order_payment aop ON aop.order_id=p_order_id
    
    WHERE 
       opl.parent_order_id = p_order_id OR opl.order_id=p_order_id ORDER BY opl.start_date ASC ;
END