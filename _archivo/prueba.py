numero1 = input("ingrese el primer numero")
numero2 = input("ingrese el segundo numero")
operacion = input("ingrese la operacion a realizar")
def suma(a, b):
    return a + b

def resta(a, b):
    return a - b

def multiplicacion(a, b):
    return a * b

def division(a, b):
    return a / b

if operacion == "suma":
    resultado = suma(int(numero1), int(numero2))
elif operacion == "resta":
    resultado = resta(int(numero1), int(numero2))
elif operacion == "multiplicacion":
    resultado = multiplicacion(int(numero1), int(numero2))
elif operacion == "division":
    resultado = division(int(numero1), int(numero2))
else:
    print("papi operacion no valida intrdocuzca otra cosa")

print(f"el resultado es {resultado}")


def suma(a, b):
    return a + b

def resta(a, b):
    return a - b

def multiplicacion(a, b):
    return a * b

def division(a, b):
    return a / b
