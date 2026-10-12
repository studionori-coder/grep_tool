package jp.example.order;

import java.math.BigDecimal;
import java.util.List;

/**
 * 受注サービス
 * 在庫引当処理を呼び出す。
 */
public class OrderService {

    private final StockService stockService;

    public OrderService(StockService stockService) {
        this.stockService = stockService;
    }

    /**
     * 受注金額を計算する。
     * @param order 受注
     * @return 金額
     */
    public BigDecimal calcOrderAmount(Order order) {
        return order.getUnitPrice().multiply(BigDecimal.valueOf(order.getQty()));
    }
}
