numero1 = input("ingrese el primer numero")
numero2 = input("ingrese el segundo numero")
operacion = input("ingrese la operacion a realizar")

if operacion == "suma":
    resultado = int(numero1) + int(numero2)
elif operacion == "resta":
    resultado = int(numero1) - int(numero2)
elif operacion == "multiplicacion":
    resultado = int(numero1) * int(numero2)
elif operacion == "division":
    resultado = int(numero1) / int(numero2)
else:
    print("papi operacion no valida intrdocuzca otra cosa")

print(f"el resultado es {resultado}")




