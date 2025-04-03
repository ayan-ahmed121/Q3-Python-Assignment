import streamlit as st
import pint

# Create a UnitRegistry
ureg = pint.UnitRegistry()

# Unit category mapping
unit_category_map = {
    "Area": "area",
    "Data Transfer Rate": "data_rate",
    "Digital Storage": "digital",
    "Energy": "energy",
    "Frequency": "frequency",
    "Fuel Economy": "fuel_efficiency",
    "Length": "length",
    "Mass": "mass",
    "Plane Angle": "angle",
    "Pressure": "pressure",
    "Speed": "speed",
    "Temperature": "temperature",
    "Time": "time",
    "Volume": "volume",
}

# Define measurement categories and their corresponding unit names
measurement_categories = list(unit_category_map.keys())

unit_type = {
    "Area": [
        "square_kilometer",
        "square_meter",
        "square_foot",
        "square_inch",
        "square_mile",
        "square_yard",
        "acre",
        "hectare",
    ],
    "Data Transfer Rate": [
        "bit / second",
        "kilobit / second",
        "kilobyte / second",
        "kibibit / second",
        "megabit / second",
        "megabyte / second",
        "mebibit / second",
        "gigabit / second",
        "gigabyte / second",
        "gibibit / second",
        "terabit / second",
        "terabyte / second",
        "tebibit / second",
    ],
    "Digital Storage": [
        "bit",
        "kilobit",
        "kibibit",
        "megabit",
        "gigabit",
        "gibibit",
        "terabit",
        "tebibit",
        "pebibit",
        "byte",
        "kilobyte",
        "kibibyte",
        "megabyte",
        "gigabyte",
        "gibibyte",
        "terabyte",
        "tebibyte",
        "petabyte",
        "pebibyte",
    ],
    "Energy": [
        "joule",
        "kilojoule",
        "calorie",
        "kilocalorie",
        "watt_hour",
        "kilowatt_hour",
        "electronvolt",
        "btu",
        "therm",
        "foot_pound",
    ],
    "Frequency": [
        "hertz", 
        "kilohertz", 
        "megahertz", 
        "gigahertz"
    ],
    "Fuel Economy": [
        "mile / gallon",
        "mile / us_gallon",
        "kilometer / liter",
        "liter / 100_kilometer",
    ],
    "Length": [
        "kilometer",
        "meter",
        "centimeter",
        "millimeter",
        "micrometer",
        "nanometer",
        "mile",
        "yard",
        "foot",
        "inch",
        "nautical_mile",
    ],
    "Mass": [
        "tonne",
        "kilogram",
        "gram",
        "milligram",
        "microgram",
        "imperial_ton",
        "us_ton",
        "stone",
        "pound",
        "ounce",
    ],
    "Plane Angle": [
        "arcsecond",
        "degree",
        "gradian",
        "milliradian",
        "arcminute",
        "radian",
    ],
    "Pressure": ["bar", "pascal", "psi", "atmosphere", "torr"],
    "Speed": [
        "mile / hour",
        "foot / second",
        "meter / second",
        "kilometer / hour",
        "knot",
    ],
    "Temperature": [
        "degree_Celsius", 
        "degree_Fahrenheit", 
        "kelvin"
    ],
    "Time":[
        "nanosecond",
        "microsecond",
        "millisecond",
        "second",
        "minute", 
        "hour",
        "day",
        "week",
        "year",
        "decade", # Define on-the-fly
        "century"
    ],
    "Volume": [
        "us_gallon",
        "us_quart",
        "us_pint",
        "us_cup",
        "us_fluid_ounce",
        "us_tablespoon",
        "us_teaspoon",
        "cubic_meter",
        "liter",
        "milliliter",
        "imperial_gallon",
        "imperial_quart",
        "imperial_pint",
        "imperial_cup",
        "imperial_fluid_ounce",
        "imperial_tablespoon",
        "imperial_teaspoon",
        "cubic_foot",
        "cubic_inch",
    ],
}



ureg.define("square_meter = meter ** 2 = sq_m")
ureg.define("square_kilometer = 1e6 * square_meter = sq_km")
ureg.define("square_yard = 0.836127 * square_meter = sq_yd")
ureg.define("kilobit_per_second = 1000 * bit / second")
ureg.define("kibibit_per_second = 1024 * bit / second")
ureg.define("mebibit_per_second = 1024 ** 2 * bit / second")
ureg.define("gibibit_per_second = 1024 ** 3 * bit / second")
ureg.define("tebibit_per_second = 1024 ** 4 * bit / second")
ureg.define("kibibit = 1024 * bit")
ureg.define("mebibit = 1024 ** 2 * bit")
ureg.define("gibibit = 1024 ** 3 * bit")
ureg.define("tebibit = 1024 ** 4 * bit")
ureg.define("pebibit = 1024 ** 5 * bit")
ureg.define("kibibyte = 1024 * byte")
ureg.define("mebibyte = 1024 ** 2 * byte")
ureg.define("gibibyte = 1024 ** 3 * byte")
ureg.define("tebibyte = 1024 ** 4 * byte")
ureg.define("pebibyte = 1024 ** 5 * byte")
ureg.define("gram_calorie = 4.184 * joule")  # 1 calorie (small calorie) = 4.184 joules
ureg.define("kilocalorie = 4184 * joule")  # 1 kilocalorie (food calorie) = 4184 joules
ureg.define("watt_hour = 3600 * joule")  # 1 Wh = 3600 joules
ureg.define("kilowatt_hour = 3.6e6 * joule")  # 1 kWh = 3.6e6 joules
ureg.define("electronvolt = 1.602176634e-19 * joule = eV")
ureg.define("btu = 1055.06 * joule")
ureg.define("us_therm = 105505585.257 * joule")  # US therm (approximate)
ureg.define("foot_pound = 1.3558179483314004 * joule")  # 1 foot-pound ≈ 1.35582 joules
ureg.define("us_gallon = 3.785411784 * liter")
ureg.define("imperial_gallon = 4.54609 * liter")
ureg.define("mile_per_us_gallon = mile / us_gallon")
ureg.define("mile_per_imperial_gallon = mile / imperial_gallon")
ureg.define("kilometer_per_liter = kilometer / liter")
ureg.define("liter_per_100km = 0.01 * liter / kilometer")
ureg.define("arcminute = 1/60 * degree")
ureg.define("arcsecond = 1/60 * arcminute")
ureg.define("gradian = 0.9 * degree")  # 1 gradian = 0.9 degrees
ureg.define("milliradian = 1e-3 * radian")
ureg.define("standard_atmosphere = 101325 * pascal")
ureg.define("torr = 133.322368 * pascal")
ureg.define("us_gallon = 3.785411784 * liter")
ureg.define("us_quart = us_gallon / 4")
ureg.define("us_pint = us_quart / 2")
ureg.define("us_fluid_ounce = us_gallon / 128")
ureg.define("us_cup = 8 * us_fluid_ounce")
ureg.define("us_tablespoon = us_fluid_ounce / 2")
ureg.define("us_teaspoon = us_tablespoon / 3")
ureg.define("us_ton = 907.18474 * kilogram") 
ureg.define("imperial_gallon = 4.54609 * liter")
ureg.define("imperial_quart = imperial_gallon / 4")
ureg.define('imperial_ton = 1016.0469088 * kilogram')
ureg.define("imperial_pint = imperial_quart / 2")
ureg.define("imperial_cup = imperial_gallon / 10")  # Note: definitions for a cup can vary
ureg.define("imperial_fluid_ounce = imperial_gallon / 160")
ureg.define("imperial_tablespoon = imperial_fluid_ounce / 2")
ureg.define("imperial_teaspoon = imperial_tablespoon / 3")
ureg.define("cubic_meter = meter ** 3")
ureg.define("cubic_foot = foot ** 3")
ureg.define("cubic_inch = inch ** 3")


def convert_temperature(value, from_unit, to_unit):
    unit_map = {
        "degree_Celsius": "Celsius",
        "degree_Fahrenheit": "Fahrenheit",
        "kelvin": "Kelvin"
        }
    from_unit = unit_map.get(from_unit, from_unit)
    to_unit = unit_map.get(to_unit, to_unit)
    
    if from_unit == "Celsius" and to_unit == "Fahrenheit":
        result = (value * 9/5) + 32
        formula = f"({value} × 9/5) + 32 = {result}"
    elif from_unit == "Fahrenheit" and to_unit == "Celsius":
        result = (value - 32) * 5/9
        formula = f"({value} - 32) × 5/9 = {result}"
    elif from_unit == "Celsius" and to_unit == "Kelvin":
        result = value + 273.15
        formula = f"{value} + 273.15 = {result}"
    elif from_unit == "Kelvin" and to_unit == "Celsius":
        result = value - 273.15
        formula = f"{value} - 273.15 = {result}"
    elif from_unit == "Fahrenheit" and to_unit == "Kelvin":
        result = (value - 32) * 5/9 + 273.15
        formula = f"(({value} - 32) × 5/9) + 273.15 = {result})"
    elif from_unit == "Kelvin" and to_unit == "Fahrenheit":
        result = (value - 273.15) * 9/5 + 32
        formula = f"(({value} - 273.15) × 9/5) + 32 = {result})"
    else:
        result = value
        formula = f"{value} {from_unit} = {value} {to_unit}"  # No conversion needed

    return result, formula

# Function for unit conversion
def convert_units(value, from_unit, to_unit, category):
    try:
        if category == "Temperature":
            return convert_temperature(value, from_unit, to_unit)
           
        elif category == "Time":
            
            time_units = {
            "second": "second",
            "minute": "minute", 
            "hour": "hour",
            "day": "day",
            "week": "week",
            "year": "year",
            "decade": "10* year", # Define on-the-fly
            "century":"100 * year",  # Define on-the-fly
            "nanosecond": "nanosecond",
            "microsecond": "microsecond",
            "millisecond": "millisecond"
        }
            # Use the unit definitions from our dictionary
            from_unit = time_units[from_unit]
            to_unit = time_units[to_unit]
        
            result = (value * ureg(from_unit)).to(ureg(to_unit))
            conversion_factor = result.magnitude / value if value != 0 else "N/A"
            formula = f"{value} {from_unit} × {conversion_factor} = {result.magnitude} {to_unit}"
            return result.magnitude, formula
        else:
            # General case for other conversions
            quantity = ureg.Quantity(value, from_unit)
            result = quantity.to(to_unit)
            conversion_factor = result.magnitude / value if value != 0 else "N/A"
            formula = f"{value} {from_unit} × {conversion_factor} = {result.magnitude} {to_unit}"
            return result.magnitude, formula
    except Exception as e:
        return None, str(e)

st.title('Unit Converter')
selected_category = st.selectbox("Select Measurement Type", measurement_categories)

from_unit = st.selectbox("From Unit:", unit_type[selected_category])
to_unit = st.selectbox("To Unit:", unit_type[selected_category])


value = st.number_input("Value:", min_value=1.0, format="%.2f")

if from_unit == to_unit:
        st.success(f"{value} {from_unit} = {value} {to_unit}")
else:
    result, formula = convert_units(value, from_unit, to_unit, selected_category)
    if result is not None:
        st.success(f"{value} {from_unit} = {result:.4f} {to_unit}")
        st.subheader('Formula')
        st.latex(formula)  # Display formula in LaTeX format
    else:
        st.error(f"Conversion error: {formula}")

