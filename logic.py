def compare_forecast_methods(series: np.ndarray) -> dict:
    """Compara todos los métodos de pronóstico y selecciona el mejor."""
    if len(series) < 3:
        return {
            "best_method": "promedio",
            "best_params": {},
            "forecast": np.mean(series) if len(series) > 0 else 0,
            "metrics": {"promedio": {"mape": 0, "rmse": 0, "mae": 0}}
        }
    
    n = len(series)
    train_size = int(n * 0.8)
    train = series[:train_size]
    test = series[train_size:]
    
    if len(test) < 1:
        train = series
        test = series[-2:]
    
    methods = {}
    
    # 1. Promedio
    forecast_mean = np.mean(train)
    methods["promedio"] = {
        "forecast": forecast_mean,
        "params": {},
        "mape": _calculate_mape(test, forecast_mean),
        "rmse": _calculate_rmse(test, forecast_mean),
        "mae": _calculate_mae(test, forecast_mean)
    }
    
    # 2. SES
    best_ses = {"mape": float("inf"), "params": {}, "forecast": 0}
    for alpha in [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
        forecast_ses = _forecast_ses(train, alpha)
        mape = _calculate_mape(test, forecast_ses)
        if mape < best_ses["mape"]:
            best_ses = {
                "mape": mape,
                "params": {"alpha": alpha},
                "forecast": forecast_ses,
                "rmse": _calculate_rmse(test, forecast_ses),
                "mae": _calculate_mae(test, forecast_ses)
            }
    methods["ses"] = best_ses
    
    # 3. Holt
    best_holt = {"mape": float("inf"), "params": {}, "forecast": 0}
    for alpha in [0.1, 0.2, 0.3, 0.4, 0.5]:
        for beta in [0.05, 0.1, 0.2, 0.3]:
            try:
                forecast_holt = _forecast_holt(train, alpha, beta)
                mape = _calculate_mape(test, forecast_holt)
                if mape < best_holt["mape"]:
                    best_holt = {
                        "mape": mape,
                        "params": {"alpha": alpha, "beta": beta},
                        "forecast": forecast_holt,
                        "rmse": _calculate_rmse(test, forecast_holt),
                        "mae": _calculate_mae(test, forecast_holt)
                    }
            except:
                continue
    methods["holt"] = best_holt
    
    # 4. Holt-Winters
    if len(series) >= 12:
        best_hw = {"mape": float("inf"), "params": {}, "forecast": 0}
        for alpha in [0.1, 0.2, 0.3, 0.4, 0.5]:
            for beta in [0.05, 0.1, 0.2, 0.3]:
                for gamma in [0.05, 0.1, 0.2]:
                    for seasonality in [12, 6, 3]:
                        try:
                            forecast_hw = _forecast_hw(train, alpha, beta, gamma, seasonality)
                            mape = _calculate_mape(test, forecast_hw)
                            if mape < best_hw["mape"]:
                                best_hw = {
                                    "mape": mape,
                                    "params": {
                                        "alpha": alpha, "beta": beta,
                                        "gamma": gamma, "seasonality": seasonality
                                    },
                                    "forecast": forecast_hw,
                                    "rmse": _calculate_rmse(test, forecast_hw),
                                    "mae": _calculate_mae(test, forecast_hw)
                                }
                        except:
                            continue
        methods["hw"] = best_hw
    
    best_method = min(methods.keys(), key=lambda m: methods[m]["mape"])
    
    return {
        "best_method": best_method,
        "best_params": methods[best_method]["params"],
        "forecast": methods[best_method]["forecast"],
        "metrics": {
            m: {
                "mape": methods[m]["mape"],
                "rmse": methods[m]["rmse"],
                "mae": methods[m]["mae"]
            }
            for m in methods
        }
    }
