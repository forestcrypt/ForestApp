def gustomota(row_data, config):
    """
    Формула из ячейки AE16:
    =SUM(...)/($AE$16*$AE$16*3.15/10000*1000)
    """
    total = sum([
        float(row_data['do_0.5m']),
        float(row_data['0.5-1.5m']),
        float(row_data['bolee_1.5m'])
    ])
    area = config.get('plot_area', 1)
    return total / (area ** 2 * 3.15 / 10000 * 1000)

def calculate_avg_height(row_data):
    # Формула из столбца H выше 1.5
    return sum(row_data['heights']) / len(row_data['heights']) if row_data['heights'] else 0