def caudal(qmax, pr, pwf):
    """
      Calcula el caudal del petroleo.

      Parámetros:
      - qmax: caudal maximo del petroleo en stb/d
      - pr: presion del reservorio en psi
      - pwf: presion de fondo fluyente en psi


      Retorna:
      - q0: caudal del petroleo en stb/d
      """
    q0 = qmax * (1 - 0.2 * (pwf / pr) - 0.8 * ((pwf / pr) ** 2))


    return q0