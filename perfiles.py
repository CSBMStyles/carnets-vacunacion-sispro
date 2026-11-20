"""Declaración de campos por proyecto.

La interfaz se dibuja a partir de esta declaración. Para reutilizarla en
policía, procuraduría, rethus, etc., basta con copiar este archivo y cambiar
`PROYECTO` y los candidatos de cada campo.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Campo:
    clave: str
    etiqueta: str
    candidatos: tuple[str, ...]
    obligatorio: bool = True


CAMPOS: tuple[Campo, ...] = (
    Campo(
        clave="documento",
        etiqueta="Documento / Cédula",
        candidatos=(
            "Identificacion",
            "Identificación",
            "Numero Documento",
            "Número de documento",
            "Numero de documento",
            "Documento",
            "Cedula",
            "Cédula",
            "CC",
        ),
    ),
    Campo(
        clave="fecha",
        etiqueta="Fecha de expedición",
        candidatos=(
            "Expedicion",
            "Expedición",
            "Fecha Expedicion",
            "Fecha de expedicion",
            "Fecha Expedición",
            "Fecha de Expedición",
        ),
    ),
    Campo(
        clave="nombre",
        etiqueta="Nombres",
        candidatos=("Nombres", "Nombre"),
    ),
    Campo(
        clave="apellido",
        etiqueta="Apellidos",
        candidatos=("Apellidos", "Apellido"),
    ),
    Campo(
        clave="perfil",
        etiqueta="Perfil / Cargo",
        candidatos=("Perfil", "Cargo", "Rol"),
    ),
)

# Metadatos del proyecto. Se usarán cuando se generalice la plantilla.
PROYECTO = {
    "id": "vacunacion",
    "nombre": "Carnet Digital de Vacunación",
    "entidad": "MiVacuna - SISPRO",
    "tipo_captcha": "imagen_ia",  # imagen_ia | recaptcha_2captcha | texto
    "salida_por_perfil": True,
    "perfiles": ("AUXILIAR", "ENFERMERO", "MEDICO"),
}
