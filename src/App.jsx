import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line, Legend } from "recharts";
import { readSheet, writeRow, deleteRow as apiDeleteRow } from "./api";

// ── DATOS INICIALES ───────────────────────────────────────────────────────────
const INIT_VACANTES = [{"id":"V-001","dpto":"CANAL TELEFONICO","posicion":"Coordinador de Contact Center","fechaSolicitud":"2025-07-29","fechaAprobacion":"2025-07-29","fechaInicio":"2025-07-30","fechaLimite":"2025-09-30","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"19000","sbaMax":"20000","sbaFinal":"20000","canal":"Infojobs","quarter":"Q3-2025"},{"id":"V-002","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2025-08-28","fechaAprobacion":"2025-08-28","fechaInicio":"2025-08-28","fechaLimite":"2025-09-30","plazas":1,"estado":"Cancelada","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Viterbit","quarter":"Q3-2025"},{"id":"V-002B","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2025-10-01","fechaAprobacion":"2025-10-03","fechaInicio":"2025-10-03","fechaLimite":"2025-11-30","plazas":1,"estado":"Cancelada","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Viterbit","quarter":"Q4-2025"},{"id":"V-003","dpto":"CANAL TELEFONICO","posicion":"Teleoperador Comercial","fechaSolicitud":"2025-08-28","fechaAprobacion":"2025-08-29","fechaInicio":"2025-08-29","fechaLimite":"2025-10-06","plazas":4,"estado":"Parcialmente cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q3-2025"},{"id":"V-003B","dpto":"CANAL TELEFONICO","posicion":"Teleoperador Comercial","fechaSolicitud":"2025-10-06","fechaAprobacion":"2025-10-06","fechaInicio":"2025-10-06","fechaLimite":"2025-10-27","plazas":2,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q4-2025"},{"id":"V-004","dpto":"PRESTACIONES SALUD","posicion":"Teleoperador B2B","fechaSolicitud":"","fechaAprobacion":"2025-09-02","fechaInicio":"2025-09-02","fechaLimite":"2025-09-29","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"19000","sbaMax":"19000","sbaFinal":"19000","canal":"Infojobs","quarter":"Q3-2025"},{"id":"V-005","dpto":"PERSONAS","posicion":"Técnico Laboral","fechaSolicitud":"2025-08-28","fechaAprobacion":"2025-08-31","fechaInicio":"2025-09-01","fechaLimite":"2025-10-16","plazas":1,"estado":"Cubierta","modalidad":"Mixta","consultora":"","dificultad":"Alto","sbaMin":"25000","sbaMax":"30000","sbaFinal":"28000","canal":"Infojobs","quarter":"Q3-2025"},{"id":"V-006","dpto":"COMERCIAL","posicion":"Gestor grandes mediadores Granada","fechaSolicitud":"2025-09-19","fechaAprobacion":"2025-09-20","fechaInicio":"","fechaLimite":"","plazas":1,"estado":"Cancelada","modalidad":"Interna","consultora":"","dificultad":"Alto","sbaMin":"","sbaMax":"","sbaFinal":"","canal":"","quarter":"Q3-2025"},{"id":"V-007","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2025-12-04","fechaAprobacion":"2025-12-04","fechaInicio":"2025-12-04","fechaLimite":"2026-01-13","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Viterbit","quarter":"Q4-2025"},{"id":"V-008","dpto":"CANAL TELEFONICO","posicion":"Teleoperador Comercial Decesos","fechaSolicitud":"2026-01-08","fechaAprobacion":"2026-01-13","fechaInicio":"2026-01-13","fechaLimite":"2026-02-16","plazas":3,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q1-2026"},{"id":"V-009","dpto":"TI","posicion":"Programadores IA","fechaSolicitud":"2025-11-27","fechaAprobacion":"2025-11-27","fechaInicio":"2025-11-27","fechaLimite":"2026-01-27","plazas":2,"estado":"Parcialmente cubierta","modalidad":"Externa","consultora":"Wehunt","dificultad":"Medio","sbaMin":"","sbaMax":"33000","sbaFinal":"33000","canal":"Consultora","quarter":"Q4-2025"},{"id":"V-010","dpto":"CANAL TELEFONICO","posicion":"Teleoperador Comercial Salud","fechaSolicitud":"2026-01-08","fechaAprobacion":"2026-01-13","fechaInicio":"2026-01-13","fechaLimite":"2026-03-02","plazas":3,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q1-2026"},{"id":"V-011","dpto":"COMERCIAL","posicion":"Gestor grandes mediadores","fechaSolicitud":"2025-11-13","fechaAprobacion":"2025-11-13","fechaInicio":"2025-11-13","fechaLimite":"2026-01-12","plazas":4,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Alto","sbaMin":"30000","sbaMax":"35000","sbaFinal":"35000","canal":"Viterbit","quarter":"Q4-2025"},{"id":"V-012","dpto":"MARKETING Y CANAL DIGITAL","posicion":"Especialista Comunicación y Marketing","fechaSolicitud":"2025-12-01","fechaAprobacion":"2025-12-01","fechaInicio":"2025-12-01","fechaLimite":"2025-12-31","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"25000","sbaMax":"30000","sbaFinal":"26000","canal":"LinkedIn","quarter":"Q4-2025"},{"id":"V-013","dpto":"PRESTACIONES SALUD","posicion":"Teleoperador B2B","fechaSolicitud":"2025-11-21","fechaAprobacion":"2025-11-21","fechaInicio":"2025-11-21","fechaLimite":"2025-12-31","plazas":2,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"19000","sbaMax":"19000","sbaFinal":"19000","canal":"Infojobs","quarter":"Q4-2025"},{"id":"V-014","dpto":"PRESTACIONES SALUD","posicion":"Coordinador de Contact Center","fechaSolicitud":"2025-11-21","fechaAprobacion":"2025-11-21","fechaInicio":"2025-11-21","fechaLimite":"2025-12-31","plazas":1,"estado":"Cubierta","modalidad":"Externa","consultora":"Robert Walters","dificultad":"Medio","sbaMin":"22000","sbaMax":"22000","sbaFinal":"22000","canal":"Consultora","quarter":"Q4-2025"},{"id":"V-015","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2026-02-19","fechaAprobacion":"2026-02-19","fechaInicio":"2026-03-16","fechaLimite":"2026-03-19","plazas":3,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q1-2026"},{"id":"V-016","dpto":"CANAL TELEFONICO","posicion":"Coordinador de Contact Center","fechaSolicitud":"2026-02-19","fechaAprobacion":"2026-02-19","fechaInicio":"2026-03-23","fechaLimite":"2026-03-19","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"20000","sbaMax":"20000","sbaFinal":"20000","canal":"Viterbit","quarter":"Q1-2026"},{"id":"V-017","dpto":"ACTUARIAL","posicion":"Actuario","fechaSolicitud":"2025-11-04","fechaAprobacion":"2025-11-04","fechaInicio":"2025-11-05","fechaLimite":"2025-12-31","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Alto","sbaMin":"30000","sbaMax":"35000","sbaFinal":"32000","canal":"Viterbit","quarter":"Q4-2025"},{"id":"V-018","dpto":"PERSONAS","posicion":"Prácticas en RRHH","fechaSolicitud":"2026-02-16","fechaAprobacion":"2026-02-16","fechaInicio":"2026-02-27","fechaLimite":"2026-02-27","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"6000","sbaMax":"6000","sbaFinal":"6000","canal":"Infojobs","quarter":"Q1-2026"},{"id":"V-019","dpto":"MARKETING Y CANAL DIGITAL","posicion":"Especialista UX/UI","fechaSolicitud":"2026-04-13","fechaAprobacion":"2026-04-13","fechaInicio":"2026-04-13","fechaLimite":"2026-05-13","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"25000","sbaMax":"35000","sbaFinal":"24000","canal":"LinkedIn","quarter":"Q2-2026"},{"id":"V-020","dpto":"CANAL TELEFONICO","posicion":"Coordinador de Contact Center","fechaSolicitud":"2025-04-15","fechaAprobacion":"2025-04-15","fechaInicio":"2025-04-15","fechaLimite":"2025-05-15","plazas":1,"estado":"Cancelada","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"19000","sbaMax":"20000","sbaFinal":"","canal":"","quarter":"Q2-2025"},{"id":"V-021","dpto":"COMERCIAL","posicion":"Gestor grandes mediadores","fechaSolicitud":"2026-04-16","fechaAprobacion":"2026-04-16","fechaInicio":"2026-04-16","fechaLimite":"2026-05-15","plazas":1,"estado":"Activa","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"30000","sbaMax":"35000","sbaFinal":"","canal":"","quarter":"Q2-2026"},{"id":"V-022","dpto":"PRESTACIONES SALUD","posicion":"Teleoperador B2B","fechaSolicitud":"2026-04-22","fechaAprobacion":"2026-04-22","fechaInicio":"2026-04-22","fechaLimite":"2026-05-22","plazas":2,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"19000","sbaMax":"19000","sbaFinal":"19000","canal":"Infojobs","quarter":"Q2-2026"},{"id":"V-023","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2026-04-28","fechaAprobacion":"2026-04-28","fechaInicio":"2026-04-28","fechaLimite":"2026-05-28","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q2-2026"},{"id":"V-024","dpto":"CANAL TELEFONICO","posicion":"Teleoperador ATC - Postventa","fechaSolicitud":"2026-05-08","fechaAprobacion":"2026-05-11","fechaInicio":"2026-05-11","fechaLimite":"2026-06-11","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Bajo","sbaMin":"16500","sbaMax":"16500","sbaFinal":"16500","canal":"Infojobs","quarter":"Q2-2026"},{"id":"V-025","dpto":"MARKETING Y CANAL DIGITAL","posicion":"Especialista Comunicación y Marketing","fechaSolicitud":"2026-05-21","fechaAprobacion":"2026-05-27","fechaInicio":"2026-05-27","fechaLimite":"2026-06-27","plazas":1,"estado":"Activa","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"25000","sbaMax":"35000","sbaFinal":"","canal":"","quarter":"Q2-2026"},{"id":"V-026","dpto":"PRESTACIONES SALUD","posicion":"Teleoperador B2B","fechaSolicitud":"2026-05-26","fechaAprobacion":"2026-05-26","fechaInicio":"2026-05-26","fechaLimite":"2026-06-26","plazas":1,"estado":"Cubierta","modalidad":"Interna","consultora":"","dificultad":"Medio","sbaMin":"19000","sbaMax":"19000","sbaFinal":"19000","canal":"Infojobs","quarter":"Q2-2026"}];

const INIT_SELECCIONADOS = [{"id":"S-001","idVacante":"V-001","nombre":"Alberto","apellidos":"Gonzalez","fechaSeleccion":"2025-09-03","fechaIncorpPrev":"2025-09-22","fechaIncorpReal":"2025-09-22","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-002","idVacante":"V-002","nombre":"Yolima","apellidos":"Cruz Garzón","fechaSeleccion":"2025-09-22","fechaIncorpPrev":"2025-09-29","fechaIncorpReal":"2025-09-29","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Desistimiento tras Aceptación"},{"id":"S-003","idVacante":"V-002B","nombre":"Lorena","apellidos":"Gomez","fechaSeleccion":"2025-11-19","fechaIncorpPrev":"2025-11-26","fechaIncorpReal":"2025-11-26","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Desistimiento tras Aceptación"},{"id":"S-004","idVacante":"V-003","nombre":"Kelmatt Enrique","apellidos":"Ayanz Cuellar","fechaSeleccion":"2025-09-19","fechaIncorpPrev":"2025-10-06","fechaIncorpReal":"2025-10-06","genero":"Hombre","tipoContrato":"Indefinido","resultado":"No-show"},{"id":"S-005","idVacante":"V-003","nombre":"Agustina","apellidos":"Torrez","fechaSeleccion":"2025-09-22","fechaIncorpPrev":"2025-10-06","fechaIncorpReal":"2025-10-06","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-006","idVacante":"V-003","nombre":"Antonio","apellidos":"Serrano Melagarejo","fechaSeleccion":"2025-10-01","fechaIncorpPrev":"2025-10-06","fechaIncorpReal":"2025-10-06","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-007","idVacante":"V-003B","nombre":"Silvia","apellidos":"Molina Burgos","fechaSeleccion":"2025-10-09","fechaIncorpPrev":"2025-10-27","fechaIncorpReal":"2025-10-27","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-008","idVacante":"V-003B","nombre":"Maria del Carmen","apellidos":"Garcia Sanchez","fechaSeleccion":"2025-10-09","fechaIncorpPrev":"2025-10-27","fechaIncorpReal":"2025-10-27","genero":"Mujer","tipoContrato":"Temporal","resultado":"Incorporado"},{"id":"S-009","idVacante":"V-004","nombre":"Jessica","apellidos":"Velazquez","fechaSeleccion":"2025-09-10","fechaIncorpPrev":"2025-09-24","fechaIncorpReal":"2025-09-24","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-010","idVacante":"V-005","nombre":"","apellidos":"","fechaSeleccion":"2025-11-11","fechaIncorpPrev":"2025-11-17","fechaIncorpReal":"2025-11-17","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-011","idVacante":"V-007","nombre":"Paola","apellidos":"Leal","fechaSeleccion":"2025-12-04","fechaIncorpPrev":"2026-01-13","fechaIncorpReal":"","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Desistimiento tras Aceptación"},{"id":"S-012","idVacante":"V-008","nombre":"Paloma","apellidos":"Verdasco","fechaSeleccion":"2026-01-23","fechaIncorpPrev":"2026-02-16","fechaIncorpReal":"2026-02-16","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-013","idVacante":"V-008","nombre":"Gian Carlo","apellidos":"Jimenez","fechaSeleccion":"2025-02-05","fechaIncorpPrev":"2026-02-16","fechaIncorpReal":"2026-02-16","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-014","idVacante":"V-008","nombre":"Daniela","apellidos":"Graziani","fechaSeleccion":"2025-02-05","fechaIncorpPrev":"2026-02-16","fechaIncorpReal":"2026-02-16","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-015","idVacante":"V-010","nombre":"Jesús","apellidos":"Sanchez","fechaSeleccion":"2026-02-20","fechaIncorpPrev":"2026-03-02","fechaIncorpReal":"2026-03-02","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-016","idVacante":"V-010","nombre":"Inés","apellidos":"Lobelle","fechaSeleccion":"2026-02-20","fechaIncorpPrev":"2026-03-02","fechaIncorpReal":"2026-03-02","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-017","idVacante":"V-010","nombre":"Victoria","apellidos":"Gallardo","fechaSeleccion":"2026-02-25","fechaIncorpPrev":"2026-03-02","fechaIncorpReal":"2026-03-02","genero":"Mujer","tipoContrato":"Indefinido","resultado":"NSPP"},{"id":"S-018","idVacante":"V-009","nombre":"Antonio","apellidos":"Rodriguez","fechaSeleccion":"2025-12-17","fechaIncorpPrev":"2026-01-19","fechaIncorpReal":"2026-01-19","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-019","idVacante":"V-011","nombre":"Maria del Pinarejo","apellidos":"Ruiz-Ocaña","fechaSeleccion":"2025-12-02","fechaIncorpPrev":"2026-01-12","fechaIncorpReal":"2026-01-12","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-020","idVacante":"V-011","nombre":"Eduardo","apellidos":"Beunza","fechaSeleccion":"2025-12-26","fechaIncorpPrev":"2026-01-12","fechaIncorpReal":"2026-01-12","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-021","idVacante":"V-011","nombre":"Alberto","apellidos":"Benito","fechaSeleccion":"2025-12-17","fechaIncorpPrev":"2026-01-12","fechaIncorpReal":"2026-01-12","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-022","idVacante":"V-011","nombre":"Iñaki","apellidos":"Parras","fechaSeleccion":"2026-01-02","fechaIncorpPrev":"2026-01-12","fechaIncorpReal":"2026-01-12","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-023","idVacante":"V-012","nombre":"Abril","apellidos":"Soria","fechaSeleccion":"2026-02-09","fechaIncorpPrev":"2026-02-23","fechaIncorpReal":"2026-02-23","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-024","idVacante":"V-014","nombre":"Raquel","apellidos":"Cardenas","fechaSeleccion":"2025-12-23","fechaIncorpPrev":"2026-01-14","fechaIncorpReal":"2026-01-14","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-025","idVacante":"V-013","nombre":"Margarita","apellidos":"Roca Dopico","fechaSeleccion":"2026-02-11","fechaIncorpPrev":"2026-02-23","fechaIncorpReal":"2026-02-23","genero":"Mujer","tipoContrato":"Indefinido","resultado":"NSPP"},{"id":"S-026","idVacante":"V-013","nombre":"Ana","apellidos":"Sueiro","fechaSeleccion":"2026-02-11","fechaIncorpPrev":"2026-02-23","fechaIncorpReal":"2026-02-23","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-027","idVacante":"V-015","nombre":"Raquel","apellidos":"Morales","fechaSeleccion":"2026-03-05","fechaIncorpPrev":"2026-03-16","fechaIncorpReal":"2026-03-16","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-028","idVacante":"V-015","nombre":"Johanna","apellidos":"Rengifo","fechaSeleccion":"2026-03-06","fechaIncorpPrev":"2026-03-16","fechaIncorpReal":"2026-03-16","genero":"Mujer","tipoContrato":"Indefinido","resultado":"NSPP"},{"id":"S-029","idVacante":"V-015","nombre":"Yolanda","apellidos":"Sanchez","fechaSeleccion":"2025-03-09","fechaIncorpPrev":"2026-03-16","fechaIncorpReal":"2026-03-16","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-030","idVacante":"V-016","nombre":"Stefania","apellidos":"Simonini","fechaSeleccion":"2026-03-09","fechaIncorpPrev":"2026-03-23","fechaIncorpReal":"2026-03-23","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Desistimiento tras Aceptación"},{"id":"S-031","idVacante":"V-017","nombre":"Abrahan","apellidos":"Calderón Riccio","fechaSeleccion":"2025-12-22","fechaIncorpPrev":"2026-01-12","fechaIncorpReal":"2026-12-01","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-032","idVacante":"V-018","nombre":"Raquel","apellidos":"Rubio","fechaSeleccion":"2026-02-27","fechaIncorpPrev":"2026-03-16","fechaIncorpReal":"2026-03-16","genero":"Mujer","tipoContrato":"Prácticas","resultado":"Incorporado"},{"id":"S-033","idVacante":"V-019","nombre":"Marcos","apellidos":"Aros Camara","fechaSeleccion":"2026-06-01","fechaIncorpPrev":"2026-06-18","fechaIncorpReal":"2026-06-18","genero":"Hombre","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-034","idVacante":"V-022","nombre":"Emelina","apellidos":"Chaparro","fechaSeleccion":"2026-05-29","fechaIncorpPrev":"2026-06-09","fechaIncorpReal":"2026-06-09","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-035","idVacante":"V-022","nombre":"Ana Maria","apellidos":"Borrego","fechaSeleccion":"2026-05-29","fechaIncorpPrev":"2026-06-09","fechaIncorpReal":"2026-06-09","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Desistimiento tras Aceptación"},{"id":"S-036","idVacante":"V-023","nombre":"Tania","apellidos":"Velarde Sanchez","fechaSeleccion":"2026-05-29","fechaIncorpPrev":"2026-06-08","fechaIncorpReal":"2026-06-08","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-037","idVacante":"V-024","nombre":"Victoria","apellidos":"Doval Gago","fechaSeleccion":"2026-06-02","fechaIncorpPrev":"2026-06-08","fechaIncorpReal":"2026-06-08","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-038","idVacante":"V-026","nombre":"María","apellidos":"Morales Viciana","fechaSeleccion":"2026-06-02","fechaIncorpPrev":"2026-06-09","fechaIncorpReal":"2026-06-09","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"},{"id":"S-039","idVacante":"V-022","nombre":"Camila","apellidos":"Franco","fechaSeleccion":"2026-06-03","fechaIncorpPrev":"2026-06-09","fechaIncorpReal":"2026-06-09","genero":"Mujer","tipoContrato":"Indefinido","resultado":"Incorporado"}];

const INIT_PLATAFORMAS = [
  { id:"p1", nombre:"Viterbit", costeAnual:4256.78, año:2025, vacantesPublicadas:36, vacantescubiertas:20 },
  { id:"p2", nombre:"Infojobs", costeAnual:2994.75, año:2025, vacantesPublicadas:10, vacantesPublicadas:20 },
  { id:"p3", nombre:"LinkedIn", costeAnual:0, año:2025, vacantesPublicadas:0, vacantesPublicadas:0 },
];

const INIT_CONSULTORAS = [
  { id:"c1", idVacante:"V-009", posicion:"Programadores IA", consultora:"Wehunt", feeAcordado:6388.8, pctSBA:0.16, sbaFinal:33000, fechaFactura:"2025-12-31", estadoPago:"Pagado", notas:"" },
  { id:"c2", idVacante:"V-014", posicion:"Coordinador de Contact Center", consultora:"Robert Walters", feeAcordado:4259.2, pctSBA:0.16, sbaFinal:22000, fechaFactura:"2025-12-30", estadoPago:"Pagado", notas:"" },
];

const DEFAULT_CONFIG = {
  appNombre: "KPI Selección",
  appEmpresa: "Aura Seguros · RRHH",
  costeHoraRRHH: 14.18,
  horasBajo: 8, horasMedio: 15, horasAlto: 25,
  plantillaTotal: 141,
  masaSalarial: 3400000,
  sbaProm: 18500,
  ttfObjetivoBajo: 21, ttfObjetivoMedio: 35, ttfObjetivoAlto: 60,
  leadTimeObjetivo: 5,
  tasaNoShowMax: 5,
  tasaBajaTempMax: 10,
  tasaRotacionObj: 15,
  tasaCoberturaObj: 85,
  pctMujeresMin: 40,
  costeViterbit: 4256.78,
  costeInfojobs: 2994.75,
  costeLinkedIn: 0,
};

// ── COLORES ───────────────────────────────────────────────────────────────────
const C = {
  bg:"#f5f6fa", surface:"#ffffff", surfaceAlt:"#f0f2f7",
  border:"#e3e7f0", accent:"#1a56db", accent2:"#0ea66e",
  accent3:"#f59e0b", accent4:"#e5365a", muted:"#9aa3b8",
  text:"#1a1f36", textDim:"#4e5a78",
  shadow:"0 1px 3px rgba(26,31,54,0.08)", shadowMd:"0 4px 16px rgba(26,31,54,0.10)",
};

const QUARTERS_ALL = ["Q1-2025","Q2-2025","Q3-2025","Q4-2025","Q1-2026","Q2-2026","Q3-2026","Q4-2026"];
const ESTADOS = ["Activa","En proceso","Cubierta","Parcialmente cubierta","Cancelada"];
const MODALIDADES = ["Interna","Externa","Mixta"];
const DPTOS = ["ACTUARIAL","CANAL TELEFONICO","COMERCIAL","FINANZAS","MARKETING Y CANAL DIGITAL","OPERACIONES","PERSONAS","PRESTACIONES SALUD","TI"];
const GENEROS = ["Mujer","Hombre","No binario","Prefiero no decir"];
const CONTRATOS = ["Indefinido","Temporal","Prácticas","Freelance"];
const RESULTADOS = ["Incorporado","Desistimiento tras Aceptación","No-show","NSPP","Baja temprana"];
const CANALES = ["Infojobs","LinkedIn","Viterbit","Consultora","Referido","Web corporativa","Otro"];
const DIFICULTADES = ["Bajo","Medio","Alto"];
const ESTADOS_PAGO = ["Pendiente","Pagado","Facturado"];

// ── UTILS ─────────────────────────────────────────────────────────────────────
const pd = v => { if(!v||v==="") return null; if(v instanceof Date) return isNaN(v)?null:v; const d=new Date(v); return isNaN(d)?null:d; };
const dd = (a,b) => { const da=pd(a),db=pd(b); if(!da||!db) return null; return Math.round((db-da)/86400000); };
const avg = arr => { const n=arr.filter(v=>v!==null&&!isNaN(v)&&v>=0); return n.length?n.reduce((a,b)=>a+b,0)/n.length:null; };
const pct = (n,d) => d?Math.round((n/d)*100):0;
const fmt = n => n===null||n===undefined?"—":typeof n==="number"?n.toLocaleString("es-ES",{maximumFractionDigits:1}):n;
const fmtE = n => n===null||n===undefined?"—":Number(n).toLocaleString("es-ES",{style:"currency",currency:"EUR",maximumFractionDigits:0});

// ── KPIs CALC ─────────────────────────────────────────────────────────────────
function calcKPIs(vacantes, seleccionados, config) {
  const horas = d => d==="Alto"?config.horasAlto:d==="Medio"?config.horasMedio:config.horasBajo;

  const vacM = vacantes.map(v => {
    const sel = seleccionados.filter(s=>s.idVacante===v.id);
    const fSel = sel.map(s=>pd(s.fechaSeleccion)).filter(Boolean);
    const maxS = fSel.length?new Date(Math.max(...fSel)):null;
    const fInc = sel.map(s=>pd(s.fechaIncorpReal)).filter(Boolean);
    const maxI = fInc.length?new Date(Math.max(...fInc)):null;
    const dtf = dd(v.fechaInicio,maxS);
    const dth = dd(v.fechaInicio,maxI);
    const lt = dd(v.fechaSolicitud,v.fechaAprobacion);
    const plazo = dd(v.fechaInicio,v.fechaLimite);
    const ok = dtf!==null&&plazo!==null?dtf<=plazo:null;
    const costeRRHH = horas(v.dificultad)*config.costeHoraRRHH;
    const costePlat = (v.canal==="Viterbit"?config.costeViterbit/36:0)+(v.canal==="Infojobs"?config.costeInfojobs/10:0)+(v.canal==="LinkedIn"?config.costeLinkedIn/5:0);
    const costeTotal = costeRRHH + costePlat;
    return {...v, dtf, dth, leadTime:lt, dentroPlazo:ok, costeRRHH, costePlat, costeTotal};
  });

  const quarters = [...new Set(vacantes.map(v=>v.quarter).filter(Boolean))].sort().reverse();
  const porQ = quarters.map(q => {
    const vQ = vacM.filter(v=>v.quarter===q);
    const sQ = seleccionados.filter(s=>{const vac=vacantes.find(v=>v.id===s.idVacante);return vac&&vac.quarter===q;});
    const cub = vQ.filter(v=>v.estado==="Cubierta").length;
    const noCanc = vQ.filter(v=>v.estado!=="Cancelada").length;
    const dtfQ = vQ.map(v=>v.dtf).filter(v=>v!==null&&v>=0);
    const ltQ = vQ.map(v=>v.leadTime).filter(v=>v!==null&&v>=0);
    const incorp = sQ.filter(s=>s.resultado==="Incorporado");
    const noShow = sQ.filter(s=>s.resultado==="No-show");
    const desist = sQ.filter(s=>s.resultado==="Desistimiento tras Aceptación");
    const mujeres = sQ.filter(s=>s.genero==="Mujer");
    const dentroP = vQ.filter(v=>v.dentroPlazo===true).length;
    const conPlazo = vQ.filter(v=>v.dentroPlazo!==null).length;
    return {
      quarter:q,
      // Volumen
      vacantes:vQ.length, cubiertas:cub, canceladas:vQ.filter(v=>v.estado==="Cancelada").length,
      parciales:vQ.filter(v=>v.estado==="Parcialmente cubierta").length,
      cobertura:pct(cub,noCanc), coberturaObj:config.tasaCoberturaObj,
      plazasPedidas:vQ.reduce((a,v)=>a+(v.plazas||0),0),
      // Tiempo
      dtfProm:dtfQ.length?Math.round(avg(dtfQ)):null,
      leadTimeProm:ltQ.length?Math.round(avg(ltQ)):null, leadTimeObj:config.leadTimeObjetivo,
      pctPlazo:conPlazo?Math.round(pct(dentroP,conPlazo)):null, plazoObj:70,
      // Calidad
      seleccionados:sQ.length, incorporados:incorp.length,
      noShows:noShow.length, desistimientos:desist.length,
      tasaNoShow:sQ.length?Math.round(pct(noShow.length,sQ.length)):0, noShowObj:config.tasaNoShowMax,
      // Igualdad
      mujeres:mujeres.length, hombres:sQ.filter(s=>s.genero==="Hombre").length,
      pctMujeres:sQ.length?pct(mujeres.length,sQ.length):0, mujeresObj:config.pctMujeresMin,
    };
  });

  const totSelec = seleccionados.length;
  const totIncorp = seleccionados.filter(s=>s.resultado==="Incorporado").length;
  const totNoShow = seleccionados.filter(s=>s.resultado==="No-show").length;
  const totDesist = seleccionados.filter(s=>s.resultado==="Desistimiento tras Aceptación").length;
  const totMuj = seleccionados.filter(s=>s.genero==="Mujer").length;
  const totHom = seleccionados.filter(s=>s.genero==="Hombre").length;
  const vCub = vacantes.filter(v=>v.estado==="Cubierta").length;
  const vNoCanc = vacantes.filter(v=>v.estado!=="Cancelada").length;
  const dtfAll = vacM.map(v=>v.dtf).filter(v=>v!==null&&v>=0);

  return {
    resumen:{
      totalVacantes:vacantes.length, cubiertas:vCub, activas:vacantes.filter(v=>v.estado==="Activa"||v.estado==="En proceso").length,
      canceladas:vacantes.filter(v=>v.estado==="Cancelada").length,
      tasaCobertura:pct(vCub,vNoCanc), coberturaObj:config.tasaCoberturaObj,
      dtfProm:dtfAll.length?Math.round(avg(dtfAll)):0,
      totalSelec:totSelec, incorporados:totIncorp, noShows:totNoShow, desistimientos:totDesist,
      mujeres:totMuj, hombres:totHom, pctMujeres:pct(totMuj,totSelec), mujeresObj:config.pctMujeresMin,
      tasaNoShow:totSelec?pct(totNoShow,totSelec):0, noShowObj:config.tasaNoShowMax,
    },
    porQ, vacM, seleccionados,
    porDpto:[...new Set(vacantes.map(v=>v.dpto).filter(Boolean))].map(d=>{
      const vD=vacantes.filter(v=>v.dpto===d);
      const sD=seleccionados.filter(s=>{const vac=vacantes.find(v=>v.id===s.idVacante);return vac&&vac.dpto===d;});
      const cub=vD.filter(v=>v.estado==="Cubierta").length;
      return{dpto:d,vacantes:vD.length,cubiertas:cub,cobertura:pct(cub,vD.filter(v=>v.estado!=="Cancelada").length),seleccionados:sD.length,mujeres:sD.filter(s=>s.genero==="Mujer").length,noShows:sD.filter(s=>s.resultado==="No-show").length};
    }).sort((a,b)=>b.vacantes-a.vacantes),
  };
}

// ── COSTOS CALC ───────────────────────────────────────────────────────────────
function calcCostos(vacantes, plataformas, consultoras, config) {
  const horas = d => d==="Alto"?config.horasAlto:d==="Medio"?config.horasMedio:config.horasBajo;
  const totalPlat = plataformas.reduce((a,p)=>a+(Number(p.costeAnual)||0),0);
  const totalConsult = consultoras.reduce((a,c)=>a+(Number(c.feeAcordado)||0),0);
  const detalleVacantes = vacantes.map(v=>{
    const h = horas(v.dificultad);
    const costeRRHH = h * config.costeHoraRRHH;
    const plat = plataformas.find(p=>p.nombre===v.canal);
    const costePlatPond = plat&&plat.vacantesPublicadas>0?(plat.costeAnual/plat.vacantesPublicadas):0;
    const consult = consultoras.filter(c=>c.idVacante===v.id);
    const costeConsult = consult.reduce((a,c)=>a+(Number(c.feeAcordado)||0),0);
    const costeTotal = costeRRHH + costePlatPond + costeConsult;
    const sba = Number(v.sbaFinal)||0;
    return{...v, h, costeRRHH, costePlatPond, costeConsult, costeTotal, pctSBA:sba?costeTotal/sba:0};
  });
  const totalRRHH = detalleVacantes.reduce((a,v)=>a+v.costeRRHH,0);
  const totalGeneral = totalRRHH + totalPlat + totalConsult;
  return{totalPlat, totalConsult, totalRRHH, totalGeneral, detalleVacantes};
}

// ── EXPORT EXCEL ──────────────────────────────────────────────────────────────
function exportExcel(kpis, costos, vacantes, seleccionados, plataformas, consultoras) {
  const wb = XLSX.utils.book_new();

  // 1 · Resumen KPIs
  const r = kpis.resumen;
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["KPI","Valor","Objetivo"],
    ["Total Vacantes", r.totalVacantes, ""],
    ["Vacantes Cubiertas", r.cubiertas, ""],
    ["Tasa Cobertura", r.tasaCobertura+"%", r.coberturaObj+"%"],
    ["Vacantes Activas", r.activas, ""],
    ["Vacantes Canceladas", r.canceladas, ""],
    ["DTF prom. (días)", r.dtfProm, ""],
    ["Total Candidatos", r.totalSelec, ""],
    ["Incorporados", r.incorporados, ""],
    ["No-shows", r.noShows, "≤"+r.noShowObj+"%"],
    ["Desistimientos", r.desistimientos, ""],
    ["Mujeres seleccionadas", r.mujeres, ""],
    ["Hombres seleccionados", r.hombres, ""],
    ["% Mujeres", r.pctMujeres+"%", "≥"+r.mujeresObj+"%"],
  ]), "Resumen KPIs");

  // 2 · KPIs por Quarter
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Quarter","Vacantes","Cubiertas","% Cobertura","DTF Prom. (d)","Lead Time (d)","% Plazo","Seleccionados","Incorporados","No-shows","% Mujeres"],
    ...kpis.porQ.map(q=>[q.quarter, q.vacantes, q.cubiertas, q.cobertura+"%", q.dtfProm??"—", q.leadTimeProm??"—", q.pctPlazo!==null?q.pctPlazo+"%":"—", q.seleccionados, q.incorporados, q.noShows, q.pctMujeres+"%"])
  ]), "KPIs por Quarter");

  // 3 · Vacantes (todos los campos)
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["ID","Departamento","Posición","Estado","Modalidad","Dificultad","Canal","Quarter","Plazas","Consultora","F. Solicitud","F. Aprobación","F. Inicio","F. Límite","SBA Mínimo (€)","SBA Máximo (€)","SBA Final (€)"],
    ...vacantes.map(v=>[
      v.id, v.dpto, v.posicion, v.estado, v.modalidad, v.dificultad, v.canal, v.quarter,
      v.plazas, v.consultora||"", v.fechaSolicitud||"", v.fechaAprobacion||"",
      v.fechaInicio||"", v.fechaLimite||"",
      v.sbaMin||"", v.sbaMax||"", v.sbaFinal||""
    ])
  ]), "Vacantes");

  // 4 · Candidatos (todos los campos)
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["ID","ID Vacante","Nombre","Apellidos","Género","Tipo Contrato","Resultado","F. Selección","F. Incorp. Prevista","F. Incorp. Real"],
    ...seleccionados.map(s=>[
      s.id, s.idVacante, s.nombre||"", s.apellidos||"",
      s.genero||"", s.tipoContrato||"", s.resultado||"",
      s.fechaSeleccion||"", s.fechaIncorpPrev||"", s.fechaIncorpReal||""
    ])
  ]), "Candidatos");

  // 5 · Costos — Plataformas
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Plataforma","Año","Coste Anual (€)","Vacantes Publicadas","Coste/Vacante (€)"],
    ...plataformas.map(p=>[p.nombre, p.año, p.costeAnual, p.vacantesPublicadas||0, p.vacantesPublicadas?Math.round(p.costeAnual/p.vacantesPublicadas):0]),
    ["TOTAL","", costos.totalPlat, "", ""]
  ]), "Costos Plataformas");

  // 6 · Costos — Consultoras
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["ID Vacante","Consultora","Fee acordado (€)","% SBA","SBA Final (€)","F. Factura","Estado pago","Notas"],
    ...consultoras.map(c=>[c.idVacante, c.consultora, c.feeAcordado, Math.round(c.pctSBA*100)+"%", c.sbaFinal||"", c.fechaFactura||"", c.estadoPago||"", c.notas||""]),
    ["TOTAL","", costos.totalConsult, "", "", "", "", ""]
  ]), "Costos Consultoras");

  // 7 · Costos — Detalle por vacante
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["ID","Posición","Dificultad","H RRHH","Coste RRHH (€)","Coste Plat. (€)","Coste Consul. (€)","COSTE TOTAL (€)","SBA Final (€)","% s/SBA"],
    ...costos.detalleVacantes.map(v=>[
      v.id, v.posicion, v.dificultad, v.h,
      Math.round(v.costeRRHH), Math.round(v.costePlatPond),
      Math.round(v.costeConsult), Math.round(v.costeTotal),
      v.sbaFinal||"", v.pctSBA?(Math.round(v.pctSBA*1000)/10)+"%":"—"
    ]),
    ["","","","","","","TOTAL GENERAL", Math.round(costos.totalGeneral),"",""]
  ]), "Detalle por Vacante");

  XLSX.writeFile(wb, "KPI_Seleccion_Export.xlsx");
}

// ── UI PRIMITIVES ─────────────────────────────────────────────────────────────
const Card = ({children,style={}}) => <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:12,padding:20,boxShadow:C.shadow,...style}}>{children}</div>;
const STitle = ({children}) => <div style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:2,borderBottom:`1px solid ${C.border}`,paddingBottom:8,marginBottom:16}}>{children}</div>;
const TT = ({active,payload,label}) => { if(!active||!payload?.length) return null; return <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",fontSize:12,boxShadow:C.shadowMd}}><div style={{color:C.textDim,marginBottom:4,fontWeight:600}}>{label}</div>{payload.map((p,i)=><div key={i} style={{color:p.color,fontWeight:600}}>{p.name}: {p.value}</div>)}</div>; };
const Btn = ({children,onClick,color=C.accent,outline=false,small=false}) => <button onClick={onClick} style={{background:outline?"transparent":color,color:outline?color:"#fff",border:`1.5px solid ${color}`,padding:small?"6px 14px":"9px 20px",borderRadius:8,cursor:"pointer",fontSize:small?12:13,fontWeight:600}}>{children}</button>;

const KPICard = ({label,value,sub,color=C.accent,warn=false}) => (
  <div style={{background:C.surface,border:`1px solid ${warn?color:C.border}`,borderRadius:10,padding:"16px 20px",boxShadow:C.shadow,borderTop:`3px solid ${color}`}}>
    <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>{label}</div>
    <div style={{fontSize:26,fontWeight:700,color,fontFamily:"Georgia,serif",lineHeight:1}}>{value}</div>
    {sub&&<div style={{fontSize:12,color:warn?color:C.textDim,marginTop:4}}>{sub}</div>}
  </div>
);

const Inp = ({label,value,onChange,type="text",options,required,placeholder,note}) => (
  <div style={{display:"flex",flexDirection:"column",gap:4}}>
    <label style={{fontSize:11,fontWeight:600,color:C.textDim,textTransform:"uppercase",letterSpacing:0.5}}>{label}{required&&<span style={{color:C.accent4}}> *</span>}</label>
    {options
      ? <select value={value} onChange={e=>onChange(e.target.value)} style={{padding:"8px 10px",borderRadius:7,border:`1.5px solid ${C.border}`,background:C.surface,color:C.text,fontSize:13,outline:"none"}}>
          <option value="">— Seleccionar —</option>
          {options.map(o=><option key={o} value={o}>{o}</option>)}
        </select>
      : <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{padding:"8px 10px",borderRadius:7,border:`1.5px solid ${C.border}`,background:C.surface,color:C.text,fontSize:13,outline:"none"}}/>
    }
    {note&&<div style={{fontSize:11,color:C.muted}}>{note}</div>}
  </div>
);

// ── DASHBOARD TAB ─────────────────────────────────────────────────────────────
function TabDashboard({kpis,config,vacantes,seleccionados}) {
  const {porQ} = kpis;
  const años = ["Total",...[...new Set(vacantes.map(v=>v.quarter).filter(Boolean).map(q=>q.split("-")[1]))].sort().reverse()];
  const [añoFiltro, setAñoFiltro] = useState("Total");

  // Filtrar datos por año seleccionado
  const vacF = añoFiltro==="Total" ? vacantes : vacantes.filter(v=>v.quarter&&v.quarter.endsWith(añoFiltro));
  const selF = añoFiltro==="Total" ? seleccionados : seleccionados.filter(s=>{ const v=vacantes.find(x=>x.id===s.idVacante); return v&&v.quarter&&v.quarter.endsWith(añoFiltro); });
  const porQF = añoFiltro==="Total" ? [...porQ].sort((a,b)=>a.quarter.localeCompare(b.quarter)) : [...porQ].filter(q=>q.quarter.endsWith(añoFiltro)).sort((a,b)=>a.quarter.localeCompare(b.quarter));

  // Recalcular resumen con datos filtrados
  const vCub = vacF.filter(v=>v.estado==="Cubierta").length;
  const vNoCanc = vacF.filter(v=>v.estado!=="Cancelada").length;
  const incorp = selF.filter(s=>s.resultado==="Incorporado").length;
  const noShow = selF.filter(s=>s.resultado==="No-show").length;
  const desist = selF.filter(s=>s.resultado==="Desistimiento tras Aceptación").length;
  const mujeres = selF.filter(s=>s.genero==="Mujer").length;
  const hombres = selF.filter(s=>s.genero==="Hombre").length;
  const tasaCob = pct(vCub, vNoCanc);
  const cobOk = tasaCob>=config.tasaCoberturaObj;
  const mujOk = pct(mujeres,selF.length)>=config.pctMujeresMin;
  const nsOk = selF.length?pct(noShow,selF.length)<=config.tasaNoShowMax:true;

  // Tabla dptos
  const dptos = [...new Set(vacF.map(v=>v.dpto).filter(Boolean))];
  const tablaDpto = dptos.map(d=>{
    const vD=vacF.filter(v=>v.dpto===d);
    const sD=selF.filter(s=>{const v=vacantes.find(x=>x.id===s.idVacante);return v&&v.dpto===d;});
    const cub=vD.filter(v=>v.estado==="Cubierta").length;
    const noC=vD.filter(v=>v.estado!=="Cancelada").length;
    return{dpto:d,vacantes:vD.length,cubiertas:cub,cobertura:pct(cub,noC),seleccionados:sD.length,mujeres:sD.filter(s=>s.genero==="Mujer").length,noShows:sD.filter(s=>s.resultado==="No-show").length};
  }).sort((a,b)=>b.vacantes-a.vacantes);

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {/* Filtro año */}
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:12,color:C.muted,fontWeight:600}}>Ver:</span>
        {años.map(a=>(
          <button key={a} onClick={()=>setAñoFiltro(a)} style={{padding:"6px 16px",borderRadius:20,border:`1.5px solid ${añoFiltro===a?C.accent:C.border}`,background:añoFiltro===a?C.accent:"transparent",color:añoFiltro===a?"#fff":C.textDim,cursor:"pointer",fontSize:13,fontWeight:600}}>
            {a}
          </button>
        ))}
      </div>

      {/* KPI Cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(148px,1fr))",gap:12}}>
        <KPICard label="Total Vacantes" value={vacF.length} color={C.accent}/>
        <KPICard label="Cubiertas" value={vCub} sub={`${tasaCob}% · obj ≥${config.tasaCoberturaObj}%`} color={cobOk?C.accent2:C.accent4} warn={!cobOk}/>
        <KPICard label="Activas" value={vacF.filter(v=>v.estado==="Activa"||v.estado==="En proceso").length} color={C.accent3}/>
        <KPICard label="Candidatos" value={selF.length} color={C.textDim}/>
        <KPICard label="Incorporados" value={incorp} color={C.accent2}/>
        <KPICard label="No-shows" value={noShow} sub={`${selF.length?pct(noShow,selF.length):0}% · obj ≤${config.tasaNoShowMax}%`} color={nsOk?C.accent2:C.accent4} warn={!nsOk}/>
        <KPICard label="Mujeres" value={mujeres} sub={`${selF.length?pct(mujeres,selF.length):0}% · obj ≥${config.pctMujeresMin}%`} color={mujOk?C.accent4:C.accent4} warn={!mujOk}/>
        <KPICard label="Hombres" value={hombres} sub={`${selF.length?pct(hombres,selF.length):0}% del total`} color={C.accent}/>
      </div>

      {/* Gráficos */}
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
        <Card>
          <STitle>Vacantes y cobertura {añoFiltro==="Total"?"— todos los quarters":añoFiltro}</STitle>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={porQF} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border}/>
              <XAxis dataKey="quarter" tick={{fill:C.muted,fontSize:11}}/>
              <YAxis tick={{fill:C.muted,fontSize:11}}/>
              <Tooltip content={<TT/>}/>
              <Legend wrapperStyle={{fontSize:12,color:C.muted}}/>
              <Bar dataKey="vacantes" name="Vacantes" fill={C.accent} radius={[4,4,0,0]}/>
              <Bar dataKey="cubiertas" name="Cubiertas" fill={C.accent2} radius={[4,4,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <STitle>Resultado incorporación</STitle>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={[{name:"Incorporados",value:incorp},{name:"Desistimientos",value:desist},{name:"No-shows",value:noShow}].filter(d=>d.value>0)} cx="50%" cy="50%" innerRadius={50} outerRadius={78} dataKey="value" paddingAngle={4}>
                {[C.accent2,C.accent3,C.accent4].map((c,i)=><Cell key={i} fill={c}/>)}
              </Pie>
              <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:12,color:C.muted}}/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"3fr 2fr",gap:16}}>
        {/* Tabla departamentos */}
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{padding:"14px 20px",borderBottom:`1px solid ${C.border}`}}><STitle>Vacantes por departamento</STitle></div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:C.surfaceAlt}}>
                {["Departamento","Vacantes","Cubiertas","% Cob.","Candidatos","Mujeres","No-shows"].map(h=>(
                  <th key={h} style={{textAlign:"left",padding:"8px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tablaDpto.map((d,i)=>(
                <tr key={d.dpto} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                  <td style={{padding:"9px 14px",color:C.text,fontWeight:500,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.dpto}</td>
                  <td style={{padding:"9px 14px",textAlign:"center",fontWeight:700,color:C.accent}}>{d.vacantes}</td>
                  <td style={{padding:"9px 14px",textAlign:"center",fontWeight:700,color:C.accent2}}>{d.cubiertas}</td>
                  <td style={{padding:"9px 14px",textAlign:"center"}}>
                    <span style={{background:d.cobertura>=config.tasaCoberturaObj?C.accent2+"18":C.accent4+"18",color:d.cobertura>=config.tasaCoberturaObj?C.accent2:C.accent4,padding:"2px 8px",borderRadius:10,fontSize:11,fontWeight:700}}>{d.cobertura}%</span>
                  </td>
                  <td style={{padding:"9px 14px",textAlign:"center",color:C.textDim}}>{d.seleccionados}</td>
                  <td style={{padding:"9px 14px",textAlign:"center",color:C.accent4}}>{d.mujeres}</td>
                  <td style={{padding:"9px 14px",textAlign:"center",color:d.noShows>0?C.accent4:C.muted}}>{d.noShows||"—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>

        {/* Contraste H/M */}
        <Card>
          <STitle>Seleccionados — Hombre vs Mujer</STitle>
          <div style={{display:"flex",flexDirection:"column",gap:16,marginTop:8}}>
            {/* Barra comparativa */}
            <div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,fontWeight:600,marginBottom:6}}>
                <span style={{color:C.accent}}>Hombres: {hombres}</span>
                <span style={{color:C.accent4}}>Mujeres: {mujeres}</span>
              </div>
              <div style={{height:20,borderRadius:10,background:C.surfaceAlt,overflow:"hidden",display:"flex"}}>
                {selF.length>0&&<>
                  <div style={{width:`${pct(hombres,selF.length)}%`,background:C.accent,transition:"width 0.4s"}}/>
                  <div style={{width:`${pct(mujeres,selF.length)}%`,background:C.accent4,transition:"width 0.4s"}}/>
                </>}
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.muted,marginTop:4}}>
                <span>{selF.length?pct(hombres,selF.length):0}%</span>
                <span>{selF.length?pct(mujeres,selF.length):0}%</span>
              </div>
            </div>
            {/* Por resultado */}
            {[{label:"Incorporados",r:"Incorporado"},{label:"No-shows",r:"No-show"},{label:"Desistimientos",r:"Desistimiento tras Aceptación"}].map(({label,r})=>{
              const t=selF.filter(s=>s.resultado===r);
              const m=t.filter(s=>s.genero==="Mujer").length;
              const h=t.filter(s=>s.genero==="Hombre").length;
              return(
                <div key={r}>
                  <div style={{fontSize:11,color:C.muted,fontWeight:600,marginBottom:4,textTransform:"uppercase",letterSpacing:0.5}}>{label} ({t.length})</div>
                  <div style={{display:"flex",gap:12}}>
                    <span style={{fontSize:13,fontWeight:700,color:C.accent}}>♂ {h}</span>
                    <span style={{fontSize:13,fontWeight:700,color:C.accent4}}>♀ {m}</span>
                  </div>
                </div>
              );
            })}
            {/* Objetivo */}
            <div style={{background:mujOk?C.accent2+"12":C.accent4+"12",border:`1px solid ${mujOk?C.accent2:C.accent4}30`,borderRadius:8,padding:"10px 14px",textAlign:"center"}}>
              <div style={{fontSize:11,color:C.muted,marginBottom:2}}>Plan de Igualdad</div>
              <div style={{fontSize:16,fontWeight:700,color:mujOk?C.accent2:C.accent4}}>{selF.length?pct(mujeres,selF.length):0}% mujeres</div>
              <div style={{fontSize:11,color:mujOk?C.accent2:C.accent4}}>{mujOk?"✓ Objetivo cumplido":"⚠ Obj. ≥"+config.pctMujeresMin+"%"}</div>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ── KPIs POR QUARTER TAB ──────────────────────────────────────────────────────
function TabKPIsQ({kpis,config,vacantes}) {
  const {porQ} = kpis;
  // quarters más reciente primero
  const quartersDesc = [...porQ].sort((a,b)=>b.quarter.localeCompare(a.quarter));
  const [vista, setVista] = useState("quarter"); // "quarter" | "año"
  const [qSel, setQSel] = useState(quartersDesc.length?quartersDesc[0].quarter:"");
  const q = porQ.find(x=>x.quarter===qSel)||quartersDesc[0];

  // Años disponibles
  const años = [...new Set(porQ.map(x=>x.quarter.split("-")[1]))].sort().reverse();
  const [añoSel, setAñoSel] = useState(años[0]||"");

  // Quarters del año seleccionado, en orden lógico Q1→Q4
  const qsDelAño = porQ.filter(x=>x.quarter.endsWith(añoSel)).sort((a,b)=>a.quarter.localeCompare(b.quarter));

  // Selecciones externas por quarter
  const extPorQ = (qData) => vacantes.filter(v=>v.quarter===qData.quarter&&(v.modalidad==="Externa"||v.modalidad==="Mixta")).length;

  const buildRows = (qData) => [
    {grupo:"VOLUMEN",items:[
      {kpi:"KV1",label:"Vacantes registradas",val:qData.vacantes,obj:null},
      {kpi:"KV2",label:"Vacantes cubiertas",val:qData.cubiertas,obj:null},
      {kpi:"KV3",label:"Vacantes parcialmente cubiertas",val:qData.parciales,obj:null},
      {kpi:"KV4",label:"Tasa de cobertura",val:qData.cobertura+"%",obj:`≥${config.tasaCoberturaObj}%`,mode:"gte",num:qData.cobertura,objN:config.tasaCoberturaObj},
      {kpi:"KV5",label:"Plazas pedidas",val:qData.plazasPedidas,obj:null},
      {kpi:"KV6",label:"Selecciones externas / mixtas",val:extPorQ(qData),obj:null},
      {kpi:"KV8",label:"Vacantes canceladas",val:qData.canceladas,obj:null},
    ]},
    {grupo:"TIEMPO",items:[
      {kpi:"KT1",label:"Lead time aprobación prom.",val:qData.leadTimeProm!==null?qData.leadTimeProm+"d":"—",obj:`≤${config.leadTimeObjetivo}d`,mode:"lte",num:qData.leadTimeProm,objN:config.leadTimeObjetivo},
      {kpi:"KT2",label:"Days to Fill prom.",val:qData.dtfProm!==null?qData.dtfProm+"d":"—",obj:null},
      {kpi:"KT8",label:"% Procesos dentro de plazo",val:qData.pctPlazo!==null?qData.pctPlazo+"%":"—",obj:"≥70%",mode:"gte",num:qData.pctPlazo,objN:70},
    ]},
    {grupo:"CALIDAD / CANDIDATOS",items:[
      {kpi:"KC1",label:"Total seleccionados",val:qData.seleccionados,obj:null},
      {kpi:"KC2",label:"Incorporados",val:qData.incorporados,obj:null},
      {kpi:"KC6",label:"Tasa de no-show",val:qData.tasaNoShow+"%",obj:`≤${config.tasaNoShowMax}%`,mode:"lte",num:qData.tasaNoShow,objN:config.tasaNoShowMax},
      {kpi:"KC7",label:"Desistimientos",val:qData.desistimientos,obj:null},
    ]},
    {grupo:"IGUALDAD",items:[
      {kpi:"KI1",label:"Mujeres seleccionadas",val:qData.mujeres,obj:null},
      {kpi:"KI2",label:"Hombres seleccionados",val:qData.hombres,obj:null},
      {kpi:"KI3",label:"% Mujeres",val:qData.pctMujeres+"%",obj:`≥${config.pctMujeresMin}%`,mode:"gte",num:qData.pctMujeres,objN:config.pctMujeresMin},
    ]},
  ];

  const TablaKPIs = ({qData}) => {
    const rows = buildRows(qData);
    return rows.map(grupo=>(
      <Card key={grupo.grupo} style={{padding:0,overflow:"hidden"}}>
        <div style={{background:C.surfaceAlt,padding:"10px 16px",borderBottom:`1px solid ${C.border}`,fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:2}}>{grupo.grupo}</div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr style={{background:C.surfaceAlt}}>
              {["Código","Indicador","Valor","Objetivo","Estado"].map(h=><th key={h} style={{textAlign:"left",padding:"9px 16px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {grupo.items.map((row,i)=>{
              const ok=row.obj&&row.num!==null?(row.mode==="gte"?row.num>=row.objN:row.num<=row.objN):null;
              return(
                <tr key={row.kpi} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                  <td style={{padding:"10px 16px",color:C.accent,fontWeight:700,fontSize:12}}>{row.kpi}</td>
                  <td style={{padding:"10px 16px",color:C.text}}>{row.label}</td>
                  <td style={{padding:"10px 16px",fontWeight:700,color:C.text}}>{row.val}</td>
                  <td style={{padding:"10px 16px",color:C.muted}}>{row.obj||"—"}</td>
                  <td style={{padding:"10px 16px"}}>{ok===null?<span style={{color:C.muted}}>—</span>:ok?<span style={{color:C.accent2,fontWeight:700,fontSize:12}}>✓ OK</span>:<span style={{color:C.accent4,fontWeight:700,fontSize:12}}>⚠ Fuera</span>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    ));
  };

  // Vista comparativa por año: tabla con Q1 Q2 Q3 Q4 como columnas
  const TablaComparativa = () => {
    const grupos = buildRows(qsDelAño[0]||{vacantes:0,cubiertas:0,parciales:0,cobertura:0,plazasPedidas:0,canceladas:0,leadTimeProm:null,dtfProm:null,pctPlazo:null,seleccionados:0,incorporados:0,tasaNoShow:0,desistimientos:0,mujeres:0,hombres:0,pctMujeres:0,quarter:""});
    const getVal = (qData, kpi) => {
      for(const g of buildRows(qData||{vacantes:0,cubiertas:0,parciales:0,cobertura:0,plazasPedidas:0,canceladas:0,leadTimeProm:null,dtfProm:null,pctPlazo:null,seleccionados:0,incorporados:0,tasaNoShow:0,desistimientos:0,mujeres:0,hombres:0,pctMujeres:0,quarter:qData?.quarter||""}))
        for(const item of g.items) if(item.kpi===kpi) return {val:item.val,num:item.num,objN:item.objN,mode:item.mode,hasObj:!!item.obj};
      return {val:"—"};
    };
    return (
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        {grupos.map(grupo=>(
          <Card key={grupo.grupo} style={{padding:0,overflow:"hidden"}}>
            <div style={{background:C.surfaceAlt,padding:"10px 16px",borderBottom:`1px solid ${C.border}`,fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:2}}>{grupo.grupo}</div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr style={{background:C.surfaceAlt}}>
                    <th style={{textAlign:"left",padding:"9px 16px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`,minWidth:220}}>Indicador</th>
                    {qsDelAño.map(x=><th key={x.quarter} style={{textAlign:"center",padding:"9px 16px",color:C.accent,fontWeight:700,fontSize:12,borderBottom:`1px solid ${C.border}`,minWidth:90}}>{x.quarter}</th>)}
                    <th style={{textAlign:"left",padding:"9px 16px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>Objetivo</th>
                  </tr>
                </thead>
                <tbody>
                  {grupo.items.map((row,i)=>(
                    <tr key={row.kpi} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                      <td style={{padding:"10px 16px",color:C.text}}>
                        <span style={{fontSize:11,color:C.accent,fontWeight:700,marginRight:8}}>{row.kpi}</span>{row.label}
                      </td>
                      {qsDelAño.map(x=>{
                        const d=getVal(x,row.kpi);
                        const ok=d.hasObj&&d.num!==null?(d.mode==="gte"?d.num>=d.objN:d.num<=d.objN):null;
                        return(
                          <td key={x.quarter} style={{padding:"10px 16px",textAlign:"center",fontWeight:700,color:ok===null?C.text:ok?C.accent2:C.accent4}}>
                            {d.val}{ok!==null&&<span style={{display:"block",fontSize:10,fontWeight:600}}>{ok?"✓":"⚠"}</span>}
                          </td>
                        );
                      })}
                      <td style={{padding:"10px 16px",color:C.muted,fontSize:12}}>{row.obj||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        ))}
      </div>
    );
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {/* Selector vista */}
      <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
        <div style={{display:"flex",background:C.surfaceAlt,borderRadius:10,padding:3,gap:2}}>
          {[{id:"quarter",label:"Por Quarter"},{id:"año",label:"Por Año (comparativo)"}].map(v=>(
            <button key={v.id} onClick={()=>setVista(v.id)} style={{padding:"7px 18px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:vista===v.id?C.surface:C.surfaceAlt,color:vista===v.id?C.accent:C.muted,boxShadow:vista===v.id?C.shadow:"none",transition:"all 0.15s"}}>
              {v.label}
            </button>
          ))}
        </div>

        {vista==="quarter" && (
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {quartersDesc.map(x=>(
              <button key={x.quarter} onClick={()=>setQSel(x.quarter)} style={{padding:"7px 18px",borderRadius:20,border:`1.5px solid ${qSel===x.quarter?C.accent:C.border}`,background:qSel===x.quarter?C.accent:"transparent",color:qSel===x.quarter?"#fff":C.textDim,cursor:"pointer",fontSize:13,fontWeight:600}}>
                {x.quarter}
              </button>
            ))}
          </div>
        )}

        {vista==="año" && (
          <div style={{display:"flex",gap:8}}>
            {años.map(a=>(
              <button key={a} onClick={()=>setAñoSel(a)} style={{padding:"7px 18px",borderRadius:20,border:`1.5px solid ${añoSel===a?C.accent:C.border}`,background:añoSel===a?C.accent:"transparent",color:añoSel===a?"#fff":C.textDim,cursor:"pointer",fontSize:13,fontWeight:600}}>
                {a}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Vista por Quarter */}
      {vista==="quarter" && q && (
        <>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
            <KPICard label="Vacantes" value={q.vacantes} color={C.accent}/>
            <KPICard label="Cobertura" value={q.cobertura+"%"} sub={`obj ≥${config.tasaCoberturaObj}%`} color={q.cobertura>=config.tasaCoberturaObj?C.accent2:C.accent4} warn={q.cobertura<config.tasaCoberturaObj}/>
            <KPICard label="DTF prom." value={q.dtfProm!==null?q.dtfProm+"d":"—"} color={C.accent}/>
            <KPICard label="Sel. externas" value={extPorQ(q)} sub="Externa o Mixta" color={C.accent3}/>
            <KPICard label="No-show" value={q.tasaNoShow+"%"} sub={`obj ≤${config.tasaNoShowMax}%`} color={q.tasaNoShow<=config.tasaNoShowMax?C.accent2:C.accent4} warn={q.tasaNoShow>config.tasaNoShowMax}/>
            <KPICard label="% Mujeres" value={q.pctMujeres+"%"} sub={`obj ≥${config.pctMujeresMin}%`} color={q.pctMujeres>=config.pctMujeresMin?C.accent2:C.accent4} warn={q.pctMujeres<config.pctMujeresMin}/>
          </div>
          <TablaKPIs qData={q}/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
            <Card>
              <STitle>DTF evolución — todos los quarters</STitle>
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={[...porQ].sort((a,b)=>a.quarter.localeCompare(b.quarter)).filter(x=>x.dtfProm!==null)}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border}/>
                  <XAxis dataKey="quarter" tick={{fill:C.muted,fontSize:11}}/>
                  <YAxis tick={{fill:C.muted,fontSize:11}} unit="d"/>
                  <Tooltip content={<TT/>}/>
                  <Line type="monotone" dataKey="dtfProm" name="DTF" stroke={C.accent} strokeWidth={2.5} dot={{r:5,fill:C.accent,stroke:C.surface,strokeWidth:2}}/>
                </LineChart>
              </ResponsiveContainer>
            </Card>
            <Card>
              <STitle>% Cobertura — todos los quarters</STitle>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={[...porQ].sort((a,b)=>a.quarter.localeCompare(b.quarter))}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border}/>
                  <XAxis dataKey="quarter" tick={{fill:C.muted,fontSize:11}}/>
                  <YAxis tick={{fill:C.muted,fontSize:11}} unit="%" domain={[0,100]}/>
                  <Tooltip content={<TT/>}/>
                  <Bar dataKey="cobertura" name="% Cobertura" fill={C.accent2} radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          </div>
        </>
      )}

      {/* Vista comparativa por año */}
      {vista==="año" && (
        <>
          {qsDelAño.length===0
            ? <div style={{textAlign:"center",padding:48,color:C.muted}}>Sin datos para {añoSel}</div>
            : <>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
                  {qsDelAño.map(x=>(
                    <div key={x.quarter} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:"14px 16px",boxShadow:C.shadow,borderTop:`3px solid ${C.accent}`}}>
                      <div style={{fontSize:13,fontWeight:700,color:C.accent,marginBottom:8}}>{x.quarter}</div>
                      <div style={{fontSize:12,color:C.textDim}}>Vacantes: <b style={{color:C.text}}>{x.vacantes}</b></div>
                      <div style={{fontSize:12,color:C.textDim}}>Cobertura: <b style={{color:x.cobertura>=config.tasaCoberturaObj?C.accent2:C.accent4}}>{x.cobertura}%</b></div>
                      <div style={{fontSize:12,color:C.textDim}}>DTF: <b style={{color:C.text}}>{x.dtfProm!==null?x.dtfProm+"d":"—"}</b></div>
                      <div style={{fontSize:12,color:C.textDim}}>Sel. ext.: <b style={{color:C.accent3}}>{extPorQ(x)}</b></div>
                    </div>
                  ))}
                </div>
                <TablaComparativa/>
              </>
          }
        </>
      )}
    </div>
  );
}

// ── COSTOS TAB ────────────────────────────────────────────────────────────────
function TabCostos({vacantes, plataformas, setPlataformas, delPlat, consultoras, setConsultoras, delConsult, config}) {
  const [showFormP, setShowFormP] = useState(false);
  const [showFormC, setShowFormC] = useState(false);
  const [editP, setEditP] = useState(null);
  const [editC, setEditC] = useState(null);
  const [fp, setFp] = useState({nombre:"",costeAnual:"",año:2025,vacantesPublicadas:""});
  const [fc, setFc] = useState({idVacante:"",posicion:"",consultora:"",feeAcordado:"",pctSBA:"",sbaFinal:"",fechaFactura:"",estadoPago:"Pendiente",notas:""});

  const costos = useMemo(()=>calcCostos(vacantes,plataformas,consultoras,config),[vacantes,plataformas,consultoras,config]);

  const savePlat = () => {
    if(!fp.nombre||!fp.costeAnual) return;
    const item = {...fp, id:editP?.id||"p"+Date.now(), costeAnual:Number(fp.costeAnual), año:Number(fp.año), vacantesPublicadas:Number(fp.vacantesPublicadas)||0};
    setPlataformas(item);
    setShowFormP(false); setEditP(null); setFp({nombre:"",costeAnual:"",año:2025,vacantesPublicadas:""});
  };
  const delPlatLocal = id => delPlat(id);
  const editPlat = p => { setEditP(p); setFp({...p}); setShowFormP(true); };

  const saveConsult = () => {
    if(!fc.idVacante||!fc.consultora) return;
    const feeReal = fc.feeAcordado||(fc.sbaFinal&&fc.pctSBA?Number(fc.sbaFinal)*Number(fc.pctSBA)/100:0);
    const item = {...fc, id:editC?.id||"c"+Date.now(), feeAcordado:Number(fc.feeAcordado)||feeReal, pctSBA:Number(fc.pctSBA)/100||0, sbaFinal:Number(fc.sbaFinal)||0};
    setConsultoras(item);
    setShowFormC(false); setEditC(null); setFc({idVacante:"",posicion:"",consultora:"",feeAcordado:"",pctSBA:"",sbaFinal:"",fechaFactura:"",estadoPago:"Pendiente",notas:""});
  };
  const delConsultLocal = id => delConsult(id);
  const editConsult = c => { setEditC(c); setFc({...c,pctSBA:Math.round(c.pctSBA*100)}); setShowFormC(true); };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:24}}>
      {/* Resumen costos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <KPICard label="Coste RRHH total" value={fmtE(Math.round(costos.totalRRHH))} color={C.accent}/>
        <KPICard label="Coste plataformas" value={fmtE(Math.round(costos.totalPlat))} color={C.accent3}/>
        <KPICard label="Honorarios consultoras" value={fmtE(Math.round(costos.totalConsult))} color={C.accent4}/>
        <KPICard label="COSTE TOTAL" value={fmtE(Math.round(costos.totalGeneral))} color={C.text}/>
      </div>

      {/* A · Plataformas */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <STitle>A · Contratos anuales — Plataformas</STitle>
          {!showFormP&&<Btn onClick={()=>{setEditP(null);setFp({nombre:"",costeAnual:"",año:2025,vacantesPublicadas:""});setShowFormP(true);}} small>+ Añadir plataforma</Btn>}
        </div>
        {showFormP&&(
          <div style={{background:C.surfaceAlt,borderRadius:10,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:12}}>
            <Inp label="Plataforma" value={fp.nombre} onChange={v=>setFp(p=>({...p,nombre:v}))} placeholder="Infojobs, LinkedIn…"/>
            <Inp label="Coste Anual (€)" type="number" value={fp.costeAnual} onChange={v=>setFp(p=>({...p,costeAnual:v}))}/>
            <Inp label="Año" type="number" value={fp.año} onChange={v=>setFp(p=>({...p,año:v}))}/>
            <Inp label="Vacantes publicadas" type="number" value={fp.vacantesPublicadas} onChange={v=>setFp(p=>({...p,vacantesPublicadas:v}))}/>
            <div style={{gridColumn:"1/-1",display:"flex",gap:8,justifyContent:"flex-end"}}>
              <Btn outline color={C.muted} onClick={()=>{setShowFormP(false);setEditP(null);}}>Cancelar</Btn>
              <Btn onClick={savePlat}>Guardar</Btn>
            </div>
          </div>
        )}
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:C.surfaceAlt}}>{["Plataforma","Año","Coste Anual","Vac. Publicadas","Coste/Vacante",""].map(h=><th key={h} style={{textAlign:"left",padding:"9px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>
            {plataformas.map((p,i)=>(
              <tr key={p.id} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                <td style={{padding:"10px 14px",fontWeight:600,color:C.text}}>{p.nombre}</td>
                <td style={{padding:"10px 14px",color:C.textDim}}>{p.año}</td>
                <td style={{padding:"10px 14px",fontWeight:700,color:C.accent}}>{fmtE(p.costeAnual)}</td>
                <td style={{padding:"10px 14px",color:C.textDim}}>{p.vacantesPublicadas||0}</td>
                <td style={{padding:"10px 14px",color:C.textDim}}>{p.vacantesPublicadas?fmtE(Math.round(p.costeAnual/p.vacantesPublicadas)):"—"}</td>
                <td style={{padding:"10px 14px",display:"flex",gap:6}}>
                  <button onClick={()=>editPlat(p)} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.textDim}}>✏️</button>
                  <button onClick={()=>delPlatLocal(p.id)} style={{background:"none",border:`1px solid ${C.accent4}40`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.accent4}}>✕</button>
                </td>
              </tr>
            ))}
            <tr style={{background:C.surfaceAlt,fontWeight:700}}>
              <td colSpan={2} style={{padding:"10px 14px",color:C.textDim}}>TOTAL</td>
              <td style={{padding:"10px 14px",color:C.accent}}>{fmtE(Math.round(costos.totalPlat))}</td>
              <td colSpan={3}/>
            </tr>
          </tbody>
        </table>
      </Card>

      {/* B · Consultoras */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <STitle>B · Honorarios consultoras externas</STitle>
          {!showFormC&&<Btn onClick={()=>{setEditC(null);setFc({idVacante:"",posicion:"",consultora:"",feeAcordado:"",pctSBA:"",sbaFinal:"",fechaFactura:"",estadoPago:"Pendiente",notas:""});setShowFormC(true);}} small>+ Añadir honorario</Btn>}
        </div>
        {showFormC&&(
          <div style={{background:C.surfaceAlt,borderRadius:10,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
            <Inp label="ID Vacante *" value={fc.idVacante} onChange={v=>setFc(p=>({...p,idVacante:v}))} options={vacantes.map(v=>v.id)}/>
            <Inp label="Consultora *" value={fc.consultora} onChange={v=>setFc(p=>({...p,consultora:v}))} placeholder="Nombre consultora"/>
            <Inp label="SBA Final (€)" type="number" value={fc.sbaFinal} onChange={v=>setFc(p=>({...p,sbaFinal:v}))}/>
            <Inp label="% sobre SBA" type="number" value={fc.pctSBA} onChange={v=>setFc(p=>({...p,pctSBA:v}))} placeholder="Ej: 16"/>
            <Inp label="Fee acordado (€)" type="number" value={fc.feeAcordado} onChange={v=>setFc(p=>({...p,feeAcordado:v}))} note="Si dejas en blanco, se calcula como SBA×%"/>
            <Inp label="Estado pago" value={fc.estadoPago} onChange={v=>setFc(p=>({...p,estadoPago:v}))} options={ESTADOS_PAGO}/>
            <Inp label="Fecha Factura" type="date" value={fc.fechaFactura} onChange={v=>setFc(p=>({...p,fechaFactura:v}))}/>
            <Inp label="Notas" value={fc.notas} onChange={v=>setFc(p=>({...p,notas:v}))} placeholder="Opcional"/>
            <div style={{gridColumn:"1/-1",display:"flex",gap:8,justifyContent:"flex-end"}}>
              <Btn outline color={C.muted} onClick={()=>{setShowFormC(false);setEditC(null);}}>Cancelar</Btn>
              <Btn onClick={saveConsult}>Guardar</Btn>
            </div>
          </div>
        )}
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:C.surfaceAlt}}>{["Vacante","Consultora","Fee (€)","% SBA","F. Factura","Estado",""].map(h=><th key={h} style={{textAlign:"left",padding:"9px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>
            {consultoras.map((c,i)=>{
              const pagColor = c.estadoPago==="Pagado"?C.accent2:c.estadoPago==="Facturado"?C.accent3:C.muted;
              return(
                <tr key={c.id} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                  <td style={{padding:"10px 14px",color:C.accent,fontWeight:700}}>{c.idVacante}</td>
                  <td style={{padding:"10px 14px",color:C.text,fontWeight:500}}>{c.consultora}</td>
                  <td style={{padding:"10px 14px",fontWeight:700,color:C.text}}>{fmtE(c.feeAcordado)}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{Math.round(c.pctSBA*100)}%</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{c.fechaFactura||"—"}</td>
                  <td style={{padding:"10px 14px"}}><span style={{background:pagColor+"18",color:pagColor,padding:"3px 9px",borderRadius:12,fontSize:11,fontWeight:700}}>{c.estadoPago}</span></td>
                  <td style={{padding:"10px 14px",display:"flex",gap:6}}>
                    <button onClick={()=>editConsult(c)} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.textDim}}>✏️</button>
                    <button onClick={()=>delConsultLocal(c.id)} style={{background:"none",border:`1px solid ${C.accent4}40`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.accent4}}>✕</button>
                  </td>
                </tr>
              );
            })}
            <tr style={{background:C.surfaceAlt,fontWeight:700}}>
              <td colSpan={2} style={{padding:"10px 14px",color:C.textDim}}>TOTAL CONSULTORAS</td>
              <td style={{padding:"10px 14px",color:C.accent4}}>{fmtE(Math.round(costos.totalConsult))}</td>
              <td colSpan={4}/>
            </tr>
          </tbody>
        </table>
      </Card>

      {/* C · Detalle por vacante */}
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{padding:"14px 20px",borderBottom:`1px solid ${C.border}`}}><STitle>C · Coste interno RRHH — detalle por vacante</STitle></div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.surfaceAlt}}>{["ID","Posición","Dificultad","H RRHH","Coste RRHH","Coste Plat.","Coste Consul.","TOTAL","SBA Final","% s/SBA"].map(h=><th key={h} style={{textAlign:"left",padding:"9px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
            <tbody>
              {costos.detalleVacantes.filter(v=>v.costeTotal>0).map((v,i)=>(
                <tr key={v.id} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                  <td style={{padding:"10px 14px",color:C.accent,fontWeight:700}}>{v.id}</td>
                  <td style={{padding:"10px 14px",color:C.text,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.posicion}</td>
                  <td style={{padding:"10px 14px"}}><span style={{background:v.dificultad==="Alto"?C.accent4+"18":v.dificultad==="Medio"?C.accent3+"18":C.accent2+"18",color:v.dificultad==="Alto"?C.accent4:v.dificultad==="Medio"?C.accent3:C.accent2,padding:"2px 8px",borderRadius:10,fontSize:11,fontWeight:700}}>{v.dificultad}</span></td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{v.h}h</td>
                  <td style={{padding:"10px 14px",color:C.text}}>{fmtE(Math.round(v.costeRRHH))}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{fmtE(Math.round(v.costePlatPond))}</td>
                  <td style={{padding:"10px 14px",color:v.costeConsult>0?C.accent4:C.muted}}>{v.costeConsult>0?fmtE(Math.round(v.costeConsult)):"—"}</td>
                  <td style={{padding:"10px 14px",fontWeight:700,color:C.text}}>{fmtE(Math.round(v.costeTotal))}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{v.sbaFinal?fmtE(Number(v.sbaFinal)):"—"}</td>
                  <td style={{padding:"10px 14px",color:v.pctSBA>0.05?C.accent4:C.accent2}}>{v.pctSBA?(Math.round(v.pctSBA*1000)/10)+"%":"—"}</td>
                </tr>
              ))}
              <tr style={{background:C.surfaceAlt,fontWeight:700}}>
                <td colSpan={6} style={{padding:"10px 14px",color:C.textDim}}>COSTE TOTAL RECLUTAMIENTO</td>
                <td colSpan={2} style={{padding:"10px 14px",color:C.text,fontSize:14}}>{fmtE(Math.round(costos.totalGeneral))}</td>
                <td colSpan={2}/>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── CONFIG TAB ────────────────────────────────────────────────────────────────
function TabConfig({config, setConfig}) {
  const set = k => v => setConfig(p=>({...p,[k]:p[k]!==undefined&&typeof p[k]==="number"?(Number(v)||0):v}));
  const grupos = [
    {titulo:"Identidad de la aplicación",items:[
      {key:"appNombre",label:"Nombre de la aplicación",type:"text",note:"Ej: KPI Selección, Reclutamiento 2026…"},
      {key:"appEmpresa",label:"Empresa / subtítulo",type:"text",note:"Ej: Aura Seguros · RRHH"},
    ]},
    {titulo:"Costos internos RRHH",items:[
      {key:"costeHoraRRHH",label:"Coste hora RRHH (€/h)",type:"number",note:"Basado en convenio colectivo"},
      {key:"horasBajo",label:"Horas dedicación — Dificultad Baja",type:"number"},
      {key:"horasMedio",label:"Horas dedicación — Dificultad Media",type:"number"},
      {key:"horasAlto",label:"Horas dedicación — Dificultad Alta",type:"number"},
    ]},
    {titulo:"Plantilla y masa salarial",items:[
      {key:"plantillaTotal",label:"Plantilla total (nº empleados)",type:"number"},
      {key:"masaSalarial",label:"Masa salarial anual estimada (€)",type:"number"},
      {key:"sbaProm",label:"SBA promedio empresa (€)",type:"number"},
    ]},
    {titulo:"Objetivos de tiempo",items:[
      {key:"ttfObjetivoBajo",label:"TTF objetivo — Dificultad Baja (días)",type:"number"},
      {key:"ttfObjetivoMedio",label:"TTF objetivo — Dificultad Media (días)",type:"number"},
      {key:"ttfObjetivoAlto",label:"TTF objetivo — Dificultad Alta (días)",type:"number"},
      {key:"leadTimeObjetivo",label:"Lead time aprobación objetivo (días)",type:"number"},
    ]},
    {titulo:"Objetivos de calidad y cobertura",items:[
      {key:"tasaCoberturaObj",label:"Tasa cobertura objetivo (%)",type:"number",note:"Ej: 85 = 85%"},
      {key:"tasaNoShowMax",label:"Tasa no-show máxima (%)",type:"number"},
      {key:"tasaBajaTempMax",label:"Tasa baja temprana máxima (%)",type:"number"},
      {key:"tasaRotacionObj",label:"Tasa rotación objetivo (%)",type:"number"},
    ]},
    {titulo:"Objetivos de igualdad",items:[
      {key:"pctMujeresMin",label:"% mínimo mujeres seleccionadas",type:"number",note:"Plan de igualdad. Ej: 40 = 40%"},
    ]},
    {titulo:"Costos plataformas (contratos anuales)",items:[
      {key:"costeViterbit",label:"Coste anual Viterbit (€)",type:"number"},
      {key:"costeInfojobs",label:"Coste anual Infojobs (€)",type:"number"},
      {key:"costeLinkedIn",label:"Coste anual LinkedIn (€)",type:"number"},
    ]},
  ];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{background:C.accent+"12",border:`1px solid ${C.accent}30`,borderRadius:10,padding:"12px 16px",fontSize:13,color:C.accent,fontWeight:500}}>
        ⚙️ Los cambios aquí se aplican automáticamente a todos los cálculos de KPIs y costos en tiempo real.
      </div>
      {grupos.map(g=>(
        <Card key={g.titulo}>
          <STitle>{g.titulo}</STitle>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            {g.items.map(item=>(
              <Inp key={item.key} label={item.label} type={item.type||"text"} value={config[item.key]} onChange={set(item.key)} note={item.note}/>
            ))}
          </div>
        </Card>
      ))}
      <Btn outline color={C.accent4} onClick={()=>{if(window.confirm("¿Restaurar todos los valores por defecto?"))setConfig(DEFAULT_CONFIG);}}>↺ Restaurar valores por defecto</Btn>
    </div>
  );
}

// ── VACANTES TAB ──────────────────────────────────────────────────────────────
function TabVacantes({vacantes,saveVacante,delVacante,seleccionados}) {
  const [showForm,setShowForm] = useState(false);
  const [editing,setEditing] = useState(null);
  const [filtro,setFiltro] = useState("Todos");
  const [busqueda,setBusqueda] = useState("");
  const [sortCol,setSortCol] = useState("id");
  const [sortDir,setSortDir] = useState("asc");
  const [f,setF] = useState({id:"",dpto:"",posicion:"",fechaSolicitud:"",fechaAprobacion:"",fechaInicio:"",fechaLimite:"",plazas:1,estado:"Activa",modalidad:"Interna",consultora:"",dificultad:"Medio",sbaMin:"",sbaMax:"",sbaFinal:"",canal:"",quarter:""});
  const set = k=>v=>setF(p=>({...p,[k]:v}));
  const nextId = () => { const nums=vacantes.map(v=>parseInt(v.id.replace(/[^0-9]/g,""))).filter(n=>!isNaN(n)); return `V-${String((nums.length?Math.max(...nums):0)+1).padStart(3,"0")}`; };
  const save = () => { if(!f.posicion||!f.dpto) return alert("Completa los campos obligatorios."); const item={...f,id:f.id||nextId(),plazas:Number(f.plazas)||1}; saveVacante(item); setShowForm(false); setEditing(null); };
  const edit = v => { setEditing(v); setF({...v}); setShowForm(true); };
  const del = id => delVacante(id);
  const estadoColor = e => ({"Cubierta":C.accent2,"Activa":C.accent,"En proceso":C.accent,"Parcialmente cubierta":C.accent3,"Cancelada":C.muted}[e]||C.muted);

  const toggleSort = col => { if(sortCol===col){setSortDir(d=>d==="asc"?"desc":"asc");}else{setSortCol(col);setSortDir("asc");} };
  const sortIcon = col => sortCol===col?(sortDir==="asc"?" ↑":" ↓"):" ↕";

  const filtered = vacantes
    .filter(v=>{
      const q=busqueda.toLowerCase();
      const matchB=!q||v.id?.toLowerCase().includes(q)||v.posicion?.toLowerCase().includes(q)||v.dpto?.toLowerCase().includes(q);
      const matchE=filtro==="Todos"||v.estado===filtro;
      return matchB&&matchE;
    })
    .sort((a,b)=>{
      const va=a[sortCol]||"", vb=b[sortCol]||"";
      const cmp=typeof va==="number"?va-vb:String(va).localeCompare(String(vb));
      return sortDir==="asc"?cmp:-cmp;
    });

  const Th = ({col,label}) => (
    <th onClick={()=>toggleSort(col)} style={{textAlign:"left",padding:"10px 14px",color:sortCol===col?C.accent:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`2px solid ${C.border}`,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
      {label}{sortIcon(col)}
    </th>
  );

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <h2 style={{margin:0,fontSize:20,fontWeight:700}}>Vacantes <span style={{fontSize:14,color:C.muted,fontWeight:400}}>({vacantes.length})</span></h2>
        {!showForm&&<Btn onClick={()=>{setEditing(null);setF({id:"",dpto:"",posicion:"",fechaSolicitud:"",fechaAprobacion:"",fechaInicio:"",fechaLimite:"",plazas:1,estado:"Activa",modalidad:"Interna",consultora:"",dificultad:"Medio",sbaMin:"",sbaMax:"",sbaFinal:"",canal:"",quarter:""});setShowForm(true);}}>+ Nueva vacante</Btn>}
      </div>
      {showForm&&(
        <Card style={{borderLeft:`3px solid ${C.accent}`}}>
          <STitle>{editing?`Editando ${editing.id}`:"Nueva vacante"}</STitle>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Inp label="ID" value={f.id} onChange={set("id")} placeholder={nextId()}/>
            <Inp label="Departamento *" value={f.dpto} onChange={set("dpto")} options={DPTOS}/>
            <div style={{gridColumn:"1/-1"}}><Inp label="Posición *" value={f.posicion} onChange={set("posicion")} placeholder="Ej: Coordinador de Contact Center"/></div>
            <Inp label="Estado *" value={f.estado} onChange={set("estado")} options={ESTADOS}/>
            <Inp label="Dificultad" value={f.dificultad} onChange={set("dificultad")} options={DIFICULTADES}/>
            <Inp label="Quarter" value={f.quarter} onChange={set("quarter")} options={QUARTERS_ALL}/>
            <Inp label="Modalidad" value={f.modalidad} onChange={set("modalidad")} options={MODALIDADES}/>
            <Inp label="Canal Principal" value={f.canal} onChange={set("canal")} options={CANALES}/>
            <Inp label="Fecha Solicitud" type="date" value={f.fechaSolicitud} onChange={set("fechaSolicitud")}/>
            <Inp label="Fecha Aprobación" type="date" value={f.fechaAprobacion} onChange={set("fechaAprobacion")}/>
            <Inp label="Fecha Inicio Proceso" type="date" value={f.fechaInicio} onChange={set("fechaInicio")}/>
            <Inp label="Fecha Límite Esperada" type="date" value={f.fechaLimite} onChange={set("fechaLimite")}/>
            <Inp label="Nº Plazas" type="number" value={f.plazas} onChange={set("plazas")}/>
            <Inp label="SBA Final (€)" type="number" value={f.sbaFinal} onChange={set("sbaFinal")}/>
            <Inp label="SBA Mínimo (€)" type="number" value={f.sbaMin} onChange={set("sbaMin")}/>
            <Inp label="SBA Máximo (€)" type="number" value={f.sbaMax} onChange={set("sbaMax")}/>
            <Inp label="Consultora" value={f.consultora} onChange={set("consultora")} placeholder="Si es proceso externo"/>
            <div style={{gridColumn:"1/-1",display:"flex",gap:10,justifyContent:"flex-end"}}>
              <Btn outline color={C.muted} onClick={()=>{setShowForm(false);setEditing(null);}}>Cancelar</Btn>
              <Btn onClick={save}>Guardar vacante</Btn>
            </div>
          </div>
        </Card>
      )}
      <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
        <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="🔍 Buscar por ID, posición, departamento…" style={{padding:"7px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontSize:13,outline:"none",minWidth:260}}/>
        {["Todos",...ESTADOS].map(e=><button key={e} onClick={()=>setFiltro(e)} style={{padding:"5px 14px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:filtro===e?C.accent:C.surfaceAlt,color:filtro===e?"#fff":C.textDim}}>{e}</button>)}
        {busqueda&&<button onClick={()=>setBusqueda("")} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:8,padding:"7px 12px",cursor:"pointer",fontSize:12,color:C.muted}}>✕</button>}
      </div>
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.surfaceAlt}}>
              <Th col="id" label="ID"/>
              <Th col="dpto" label="Departamento"/>
              <Th col="posicion" label="Posición"/>
              <Th col="estado" label="Estado"/>
              <Th col="fechaInicio" label="F. Inicio"/>
              <Th col="quarter" label="Quarter"/>
              <th style={{padding:"10px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`2px solid ${C.border}`}}>DTF</th>
              <th style={{padding:"10px 14px",color:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`2px solid ${C.border}`}}>Plazo</th>
              <th style={{padding:"10px 14px",borderBottom:`2px solid ${C.border}`}}></th>
            </tr></thead>
            <tbody>
              {filtered.length===0&&<tr><td colSpan={9} style={{padding:32,textAlign:"center",color:C.muted}}>Sin vacantes con este filtro.</td></tr>}
              {filtered.map((v,i)=>{
                const sel=seleccionados.filter(s=>s.idVacante===v.id);
                const fS=sel.map(s=>pd(s.fechaSeleccion)).filter(Boolean);
                const maxS=fS.length?new Date(Math.max(...fS)):null;
                const dtf=dd(v.fechaInicio,maxS);
                const plazo=dd(v.fechaInicio,v.fechaLimite);
                const ok=dtf!==null&&plazo!==null?dtf<=plazo:null;
                const fInicioFmt = v.fechaInicio ? new Date(v.fechaInicio).toLocaleDateString("es-ES",{day:"2-digit",month:"2-digit",year:"2-digit"}) : "—";
                return(
                  <tr key={v.id} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                    <td style={{padding:"10px 14px",color:C.accent,fontWeight:700}}>{v.id}</td>
                    <td style={{padding:"10px 14px",color:C.textDim,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.dpto}</td>
                    <td style={{padding:"10px 14px",color:C.text,fontWeight:500}}>{v.posicion}</td>
                    <td style={{padding:"10px 14px"}}><span style={{background:estadoColor(v.estado)+"18",color:estadoColor(v.estado),padding:"3px 9px",borderRadius:12,fontSize:11,fontWeight:700}}>{v.estado}</span></td>
                    <td style={{padding:"10px 14px",color:C.textDim,whiteSpace:"nowrap"}}>{fInicioFmt}</td>
                    <td style={{padding:"10px 14px",color:C.textDim}}>{v.quarter||"—"}</td>
                    <td style={{padding:"10px 14px",color:C.text}}>{dtf!==null&&dtf>=0?`${dtf}d`:"—"}</td>
                    <td style={{padding:"10px 14px"}}>{ok===null?<span style={{color:C.muted}}>—</span>:ok?<span style={{color:C.accent2,fontWeight:700}}>✓</span>:<span style={{color:C.accent4,fontWeight:700}}>✗</span>}</td>
                    <td style={{padding:"10px 14px"}}><div style={{display:"flex",gap:6}}><button onClick={()=>edit(v)} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.textDim}}>✏️</button><button onClick={()=>del(v.id)} style={{background:"none",border:`1px solid ${C.accent4}40`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.accent4}}>✕</button></div></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── SELECCIONADOS TAB ─────────────────────────────────────────────────────────
function TabSeleccionados({seleccionados,saveSelec,delSelec,vacantes}) {
  const [showForm,setShowForm] = useState(false);
  const [editing,setEditing] = useState(null);
  const [f,setF] = useState({id:"",idVacante:"",nombre:"",apellidos:"",genero:"",tipoContrato:"Indefinido",resultado:"Incorporado",fechaSeleccion:"",fechaIncorpPrev:"",fechaIncorpReal:""});
  const [busqueda,setBusqueda] = useState("");
  const [filtroRes,setFiltroRes] = useState("Todos");
  const [filtroGen,setFiltroGen] = useState("Todos");
  const [sortCol,setSortCol] = useState("id");
  const [sortDir,setSortDir] = useState("asc");

  const set = k=>v=>setF(p=>({...p,[k]:v}));
  const nextId = () => { const nums=seleccionados.map(s=>parseInt(s.id.replace(/[^0-9]/g,""))).filter(n=>!isNaN(n)); return `S-${String((nums.length?Math.max(...nums):0)+1).padStart(3,"0")}`; };
  const save = () => { if(!f.idVacante||!f.nombre) return alert("Completa los campos obligatorios."); const item={...f,id:f.id||nextId()}; saveSelec(item); setShowForm(false); setEditing(null); };
  const edit = s => { setEditing(s); setF({...s}); setShowForm(true); };
  const del = id => delSelec(id);
  const resColor = r => ({"Incorporado":C.accent2,"No-show":C.accent4,"Desistimiento tras Aceptación":C.accent3,"NSPP":C.muted}[r]||C.muted);

  const exportarCandidatos = () => {
    const wb = XLSX.utils.book_new();
    const headers = ["ID","ID Vacante","Nombre","Apellidos","Género","Tipo Contrato","Resultado","F. Selección","F. Incorp. Prevista","F. Incorp. Real"];
    const rows = seleccionados.map(s=>[s.id,s.idVacante,s.nombre,s.apellidos,s.genero,s.tipoContrato,s.resultado,s.fechaSeleccion,s.fechaIncorpPrev,s.fechaIncorpReal]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers,...rows]), "Candidatos");
    XLSX.writeFile(wb, "Candidatos_Seleccionados.xlsx");
  };

  const toggleSort = col => { if(sortCol===col){setSortDir(d=>d==="asc"?"desc":"asc");}else{setSortCol(col);setSortDir("asc");} };
  const sortIcon = col => sortCol===col?(sortDir==="asc"?" ↑":" ↓"):" ↕";

  const filtered = seleccionados
    .filter(s=>{
      const q = busqueda.toLowerCase();
      const matchBusq = !q || s.nombre?.toLowerCase().includes(q)||s.apellidos?.toLowerCase().includes(q)||s.idVacante?.toLowerCase().includes(q)||s.id?.toLowerCase().includes(q);
      const matchRes = filtroRes==="Todos"||s.resultado===filtroRes;
      const matchGen = filtroGen==="Todos"||s.genero===filtroGen;
      return matchBusq&&matchRes&&matchGen;
    })
    .sort((a,b)=>{
      const va=a[sortCol]||"", vb=b[sortCol]||"";
      const cmp = typeof va==="number"?va-vb:String(va).localeCompare(String(vb));
      return sortDir==="asc"?cmp:-cmp;
    });

  const Th = ({col,label}) => (
    <th onClick={()=>toggleSort(col)} style={{textAlign:"left",padding:"10px 14px",color:sortCol===col?C.accent:C.muted,fontWeight:700,fontSize:11,textTransform:"uppercase",borderBottom:`2px solid ${C.border}`,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
      {label}{sortIcon(col)}
    </th>
  );

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <h2 style={{margin:0,fontSize:20,fontWeight:700}}>Candidatos <span style={{fontSize:14,color:C.muted,fontWeight:400}}>({filtered.length}/{seleccionados.length})</span></h2>
        <div style={{display:"flex",gap:8}}>
          <Btn onClick={exportarCandidatos} color={C.accent2} small>↓ Exportar candidatos</Btn>
          {!showForm&&<Btn onClick={()=>{if(!vacantes.length){alert("Añade primero una vacante.");return;}setEditing(null);setF({id:"",idVacante:"",nombre:"",apellidos:"",genero:"",tipoContrato:"Indefinido",resultado:"Incorporado",fechaSeleccion:"",fechaIncorpPrev:"",fechaIncorpReal:""});setShowForm(true);}}>+ Nuevo candidato</Btn>}
        </div>
      </div>

      {showForm&&(
        <Card style={{borderLeft:`3px solid ${C.accent2}`}}>
          <STitle>{editing?`Editando ${editing.id}`:"Nuevo candidato"}</STitle>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Inp label="Vacante vinculada *" value={f.idVacante} onChange={set("idVacante")} options={vacantes.map(v=>v.id)}/>
            <Inp label="Resultado *" value={f.resultado} onChange={set("resultado")} options={RESULTADOS}/>
            <Inp label="Nombre *" value={f.nombre} onChange={set("nombre")} placeholder="Nombre"/>
            <Inp label="Apellidos" value={f.apellidos} onChange={set("apellidos")} placeholder="Apellidos"/>
            <Inp label="Género" value={f.genero} onChange={set("genero")} options={GENEROS}/>
            <Inp label="Tipo de Contrato" value={f.tipoContrato} onChange={set("tipoContrato")} options={CONTRATOS}/>
            <Inp label="Fecha Selección" type="date" value={f.fechaSeleccion} onChange={set("fechaSeleccion")}/>
            <Inp label="Fecha Incorp. Prevista" type="date" value={f.fechaIncorpPrev} onChange={set("fechaIncorpPrev")}/>
            <Inp label="Fecha Incorp. Real" type="date" value={f.fechaIncorpReal} onChange={set("fechaIncorpReal")}/>
            <div style={{gridColumn:"1/-1",display:"flex",gap:10,justifyContent:"flex-end"}}>
              <Btn outline color={C.muted} onClick={()=>{setShowForm(false);setEditing(null);}}>Cancelar</Btn>
              <Btn onClick={save}>Guardar candidato</Btn>
            </div>
          </div>
        </Card>
      )}

      {/* Filtros */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
        <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="🔍 Buscar por nombre, ID vacante…" style={{padding:"7px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontSize:13,outline:"none",minWidth:220}}/>
        <select value={filtroRes} onChange={e=>setFiltroRes(e.target.value)} style={{padding:"7px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontSize:13,outline:"none",color:C.text}}>
          <option value="Todos">Todos los resultados</option>
          {RESULTADOS.map(r=><option key={r} value={r}>{r}</option>)}
        </select>
        <select value={filtroGen} onChange={e=>setFiltroGen(e.target.value)} style={{padding:"7px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontSize:13,outline:"none",color:C.text}}>
          <option value="Todos">Todos los géneros</option>
          {GENEROS.map(g=><option key={g} value={g}>{g}</option>)}
        </select>
        {(busqueda||filtroRes!=="Todos"||filtroGen!=="Todos")&&<button onClick={()=>{setBusqueda("");setFiltroRes("Todos");setFiltroGen("Todos");}} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:8,padding:"7px 12px",cursor:"pointer",fontSize:12,color:C.muted}}>✕ Limpiar</button>}
      </div>

      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:C.surfaceAlt}}>
              <Th col="id" label="ID"/>
              <Th col="idVacante" label="Vacante"/>
              <Th col="nombre" label="Nombre completo"/>
              <Th col="genero" label="Género"/>
              <Th col="tipoContrato" label="Contrato"/>
              <Th col="resultado" label="Resultado"/>
              <Th col="fechaIncorpReal" label="F. Incorp. Real"/>
              <th style={{padding:"10px 14px",borderBottom:`2px solid ${C.border}`}}></th>
            </tr></thead>
            <tbody>
              {filtered.length===0&&<tr><td colSpan={8} style={{padding:32,textAlign:"center",color:C.muted}}>Sin resultados.</td></tr>}
              {filtered.map((s,i)=>(
                <tr key={s.id} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"transparent":C.surfaceAlt+"60"}}>
                  <td style={{padding:"10px 14px",color:C.accent,fontWeight:700}}>{s.id}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{s.idVacante}</td>
                  <td style={{padding:"10px 14px",color:C.text,fontWeight:500}}>{s.nombre} {s.apellidos}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{s.genero||"—"}</td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{s.tipoContrato||"—"}</td>
                  <td style={{padding:"10px 14px"}}><span style={{background:resColor(s.resultado)+"18",color:resColor(s.resultado),padding:"3px 9px",borderRadius:12,fontSize:11,fontWeight:700}}>{s.resultado}</span></td>
                  <td style={{padding:"10px 14px",color:C.textDim}}>{s.fechaIncorpReal||"—"}</td>
                  <td style={{padding:"10px 14px"}}><div style={{display:"flex",gap:6}}><button onClick={()=>edit(s)} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.textDim}}>✏️</button><button onClick={()=>del(s.id)} style={{background:"none",border:`1px solid ${C.accent4}40`,borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:11,color:C.accent4}}>✕</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── APP ROOT ──────────────────────────────────────────────────────────────────
export default function App() {
  const [vacantes, setVacantesRaw] = useState([]);
  const [seleccionados, setSeleccionadosRaw] = useState([]);
  const [plataformas, setPlataformasRaw] = useState([]);
  const [consultoras, setConsultorasRaw] = useState([]);
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [tab, setTab] = useState("dashboard");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [saving, setSaving] = useState(false);

  // ── CARGAR DATOS DESDE GOOGLE SHEETS ──────────────────────────────────────
  const cargarDatos = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const [v, s, p, c] = await Promise.all([
        readSheet("Vacantes"),
        readSheet("Candidatos"),
        readSheet("Plataformas"),
        readSheet("Consultoras"),
      ]);
      setVacantesRaw(v.map(x=>({...x, plazas:Number(x.plazas)||1})));
      setSeleccionadosRaw(s);
      setPlataformasRaw(p.map(x=>({...x, costeAnual:Number(x.costeAnual)||0, vacantesPublicadas:Number(x.vacantesPublicadas)||0})));
      setConsultorasRaw(c.map(x=>({...x, feeAcordado:Number(x.feeAcordado)||0, pctSBA:Number(x.pctSBA)||0, sbaFinal:Number(x.sbaFinal)||0})));
    } catch(e) {
      setError("Error al conectar con Google Sheets. Comprueba tu conexión.");
    }
    setLoading(false);
  }, []);

  useEffect(() => { cargarDatos(); }, [cargarDatos]);

  // ── GUARDAR / ELIMINAR CON SYNC A SHEETS ──────────────────────────────────
  const saveVacante = useCallback(async (v) => {
    setSaving(true);
    setVacantesRaw(p => p.find(x=>x.id===v.id) ? p.map(x=>x.id===v.id?v:x) : [...p,v]);
    await writeRow("Vacantes", v);
    setSaving(false);
  }, []);

  const delVacante = useCallback(async (id) => {
    if(!window.confirm(`¿Eliminar vacante ${id}?`)) return;
    setVacantesRaw(p=>p.filter(x=>x.id!==id));
    await apiDeleteRow("Vacantes", id);
  }, []);

  const saveSelec = useCallback(async (s) => {
    setSaving(true);
    setSeleccionadosRaw(p => p.find(x=>x.id===s.id) ? p.map(x=>x.id===s.id?s:x) : [...p,s]);
    await writeRow("Candidatos", s);
    setSaving(false);
  }, []);

  const delSelec = useCallback(async (id) => {
    if(!window.confirm(`¿Eliminar candidato ${id}?`)) return;
    setSeleccionadosRaw(p=>p.filter(x=>x.id!==id));
    await apiDeleteRow("Candidatos", id);
  }, []);

  const savePlat = useCallback(async (p) => {
    setSaving(true);
    setPlataformasRaw(prev => prev.find(x=>x.id===p.id) ? prev.map(x=>x.id===p.id?p:x) : [...prev,p]);
    await writeRow("Plataformas", p);
    setSaving(false);
  }, []);

  const delPlat = useCallback(async (id) => {
    if(!window.confirm("¿Eliminar plataforma?")) return;
    setPlataformasRaw(p=>p.filter(x=>x.id!==id));
    await apiDeleteRow("Plataformas", id);
  }, []);

  const saveConsult = useCallback(async (c) => {
    setSaving(true);
    setConsultorasRaw(prev => prev.find(x=>x.id===c.id) ? prev.map(x=>x.id===c.id?c:x) : [...prev,c]);
    await writeRow("Consultoras", c);
    setSaving(false);
  }, []);

  const delConsult = useCallback(async (id) => {
    if(!window.confirm("¿Eliminar honorario?")) return;
    setConsultorasRaw(p=>p.filter(x=>x.id!==id));
    await apiDeleteRow("Consultoras", id);
  }, []);

  const kpis = useMemo(()=>vacantes.length?calcKPIs(vacantes,seleccionados,config):null,[vacantes,seleccionados,config]);
  const costos = useMemo(()=>calcCostos(vacantes,plataformas,consultoras,config),[vacantes,plataformas,consultoras,config]);

  // ── PANTALLA DE CARGA ──────────────────────────────────────────────────────
  if (loading) return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{width:40,height:40,border:`3px solid ${C.border}`,borderTop:`3px solid ${C.accent}`,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{color:C.muted,fontSize:14}}>Cargando datos desde Google Sheets…</div>
    </div>
  );

  if (error) return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{fontSize:32}}>⚠️</div>
      <div style={{color:C.text,fontSize:16}}>{error}</div>
      <button onClick={cargarDatos} style={{background:C.accent,color:"#fff",border:"none",padding:"10px 24px",borderRadius:8,cursor:"pointer",fontSize:14}}>Reintentar</button>
    </div>
  );

  const tabs = [
    {id:"dashboard",label:"📊 Dashboard"},
    {id:"porQ",label:"📈 KPIs por Quarter"},
    {id:"costos",label:"💶 Costos"},
    {id:"vacantes",label:`📋 Vacantes (${vacantes.length})`},
    {id:"seleccionados",label:`👥 Candidatos (${seleccionados.length})`},
    {id:"config",label:"⚙️ Configuración"},
  ];

  return (
    <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <style>{`*{box-sizing:border-box} input:focus,select:focus{border-color:${C.accent}!important;box-shadow:0 0 0 3px ${C.accent}18} button:hover{opacity:0.85}`}</style>

      {/* TOPBAR */}
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:58,boxShadow:C.shadow,position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:34,height:34,background:C.accent,borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17}}>📊</div>
          <div>
            <div style={{fontWeight:700,fontSize:15,color:C.text,lineHeight:1.2}}>{config.appNombre||"KPI Selección"}</div>
            <div style={{fontSize:11,color:C.muted}}>{config.appEmpresa||"RRHH"}</div>
          </div>
          <div style={{marginLeft:8,display:"flex",gap:4}}>
            <span style={{background:C.accent+"15",color:C.accent,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{vacantes.length} vacantes</span>
            <span style={{background:C.accent2+"15",color:C.accent2,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{seleccionados.length} candidatos</span>
            {saving && <span style={{background:C.accent3+"20",color:C.accent3,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>💾 Guardando…</span>}
          </div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={cargarDatos} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.muted,padding:"6px 14px",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:600}}>↺ Sincronizar</button>
          {kpis && <Btn onClick={()=>exportExcel(kpis,costos,vacantes,seleccionados,plataformas,consultoras)} color={C.accent2} small>↓ Exportar Excel</Btn>}
        </div>
      </div>

      {/* TABS */}
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"0 28px",display:"flex",gap:2,overflowX:"auto"}}>
        {tabs.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{background:"transparent",border:"none",cursor:"pointer",padding:"14px 16px",fontSize:13,fontWeight:600,color:tab===t.id?C.accent:C.muted,borderBottom:tab===t.id?`2px solid ${C.accent}`:"2px solid transparent",transition:"all 0.15s",whiteSpace:"nowrap"}}>{t.label}</button>
        ))}
      </div>

      <div style={{padding:"28px",maxWidth:1200,margin:"0 auto"}}>
        {tab==="dashboard" && (kpis
          ? <TabDashboard kpis={kpis} config={config} vacantes={vacantes} seleccionados={seleccionados}/>
          : <div style={{textAlign:"center",padding:48,color:C.muted}}>Sin datos todavía. Añade vacantes para ver el dashboard.</div>
        )}
        {tab==="porQ" && kpis && <TabKPIsQ kpis={kpis} config={config} vacantes={vacantes}/>}
        {tab==="costos" && <TabCostos vacantes={vacantes} plataformas={plataformas} setPlataformas={savePlat} delPlat={delPlat} consultoras={consultoras} setConsultoras={saveConsult} delConsult={delConsult} config={config}/>}
        {tab==="vacantes" && <TabVacantes vacantes={vacantes} saveVacante={saveVacante} delVacante={delVacante} seleccionados={seleccionados}/>}
        {tab==="seleccionados" && <TabSeleccionados seleccionados={seleccionados} saveSelec={saveSelec} delSelec={delSelec} vacantes={vacantes}/>}
        {tab==="config" && <TabConfig config={config} setConfig={setConfig}/>}
      </div>
    </div>
  );
}
