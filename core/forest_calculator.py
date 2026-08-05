import math


def calculate_plot_area(radius_m):
    return math.pi * (radius_m ** 2)


def calculate_area_ha(radius_m):
    return calculate_plot_area(radius_m) / 10000


def calculate_trees_per_ha(radius_m, tree_count):
    area_ha = calculate_area_ha(radius_m)
    return tree_count / area_ha if area_ha > 0 else 0


def calculate_density(radius_m, count):
    area_ha = calculate_area_ha(radius_m)
    return count / area_ha if area_ha > 0 else 0


def calculate_stock(avg_diameter_cm, avg_height_m, density, form_factor=0.5):
    volume_per_tree = (math.pi * (avg_diameter_cm / 200) ** 2) * avg_height_m * form_factor
    return volume_per_tree * density


def calculate_composition_coefficient(species_densities):
    total = sum(species_densities.values())
    if total == 0:
        return {s: 0 for s in species_densities}
    coeffs = {}
    for species, density in species_densities.items():
        coeffs[species] = max(1, round(density / total * 10))
    while sum(coeffs.values()) != 10:
        max_s = max(coeffs, key=coeffs.get)
        if sum(coeffs.values()) > 10:
            coeffs[max_s] -= 1
        else:
            coeffs[max_s] += 1
    return coeffs


def calculate_age_by_height(height_m, breed_type='deciduous'):
    if breed_type == 'coniferous':
        if height_m < 0.5:
            return 2
        elif height_m < 1.5:
            return 5
        else:
            return 10
    else:
        if height_m < 1.0:
            return 3
        elif height_m < 3.0:
            return 7
        else:
            return 12


def calculate_intensity(current_density, target_density):
    if current_density <= 0:
        return 0
    return max(0, (current_density - target_density) / current_density * 100)


def calculate_basal_area(diameters_cm):
    return sum(math.pi * (d / 200) ** 2 for d in diameters_cm)


def calculate_avg_height_from_dbh(height_data):
    heights = [h for h in height_data if h > 0]
    return sum(heights) / len(heights) if heights else 0


def calculate_michaelis_formula(c, b, d):
    h = d ** 2 / (c + b * d + d ** 2)
    return h * 10 if h else 0
