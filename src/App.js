import React, { useState, useEffect, useCallback } from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

// URL base de tu API Gateway en AWS
//const API_BASE_URL = "https://3fkbwpfxpk.execute-api.us-east-1.amazonaws.com/v1/solicitudes";
const API_BASE_URL = "http://localhost:8080/v1/solicitudes";

function App() {
    const { instance, accounts } = useMsal();
    const [solicitudes, setSolicitudes] = useState([]);
    const [tipoSoporte, setTipoSoporte] = useState("");

    const handleLogin = () => {
        instance.loginRedirect(loginRequest).catch(e => console.error(e));
    };

    const handleLogout = () => {
        instance.logoutRedirect().catch(e => console.error(e));
    };

    // Función auxiliar para obtener el token silenciosamente
    const getToken = async () => {
        try {
            const response = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            return response.accessToken;
        } catch (error) {
            console.error("Error al obtener el token: ", error);
            throw error;
        }
    };

    // 1. OBTENER SOLICITUDES (GET)
    const fetchSolicitudes = useCallback(async () => {
        try {
            const token = await getToken();
            const response = await fetch(API_BASE_URL, {
                headers: { "Authorization": `Bearer ${token}` }
            });
            if (!response.ok) throw new Error("Error fetching");
            const data = await response.json();
            setSolicitudes(data);
        } catch (error) {
            console.error("Error cargando solicitudes", error);
        }
    }, [accounts, instance]);

    // Cargar las solicitudes apenas el usuario inicia sesión
    useEffect(() => {
        if (accounts.length > 0) {
            fetchSolicitudes();
        }
    }, [accounts, fetchSolicitudes]);

    // 2. CREAR SOLICITUD (POST)
    const crearSolicitud = async (e) => {
        e.preventDefault();
        try {
            const token = await getToken();
            const nuevaSolicitud = {
                usuario: accounts[0].username, // Usamos el correo del usuario logueado
                tipoSoporte: tipoSoporte
                // El backend se encarga de poner estado "CREADA"
            };

            const response = await fetch(API_BASE_URL, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(nuevaSolicitud)
            });

            if (response.ok) {
                setTipoSoporte(""); // Limpiar input
                fetchSolicitudes(); // Recargar tabla
                alert("Solicitud creada exitosamente");
            } else {
                alert("Error al crear la solicitud");
            }
        } catch (error) {
            console.error("Error creando", error);
        }
    };

    // 3. CAMBIAR ESTADO DE LA SOLICITUD (PUT)
    const cambiarEstado = async (id, nuevoEstado) => {
        try {
            const token = await getToken();
            // Asumiendo que tu BFF redirige esto a /api/solicitudes/{id}/estado en el microservicio
            const response = await fetch(`${API_BASE_URL}/${id}/estado`, {
                method: "PUT",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ estado: nuevoEstado })
            });

            if (response.ok) {
                fetchSolicitudes(); // Recargar tabla para ver el cambio
            } else {
                const errMsg = await response.text();
                alert("Error al cambiar estado: " + errMsg);
            }
        } catch (error) {
            console.error("Error actualizando", error);
        }
    };

    return (
        <div style={{ padding: "30px", fontFamily: "sans-serif", maxWidth: "900px", margin: "0 auto" }}>
            <h1>MesaTech Cloud - MVP</h1>

            <UnauthenticatedTemplate>
                <div style={{ border: "1px solid #ccc", padding: "20px", borderRadius: "5px", textAlign: "center" }}>
                    <h2>Acceso Restringido</h2>
                    <p>Inicia sesión con Microsoft Entra ID para acceder al sistema.</p>
                    <button onClick={handleLogin} style={{ padding: "10px 20px", backgroundColor: "#0078D4", color: "white", border: "none", cursor: "pointer", fontSize: "16px" }}>
                        Iniciar Sesión
                    </button>
                </div>
            </UnauthenticatedTemplate>

            <AuthenticatedTemplate>
                <div style={{ border: "1px solid #28a745", padding: "15px", borderRadius: "5px", marginBottom: "20px", backgroundColor: "#f9fff9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div>
                        <strong>Usuario:</strong> {accounts[0]?.name} ({accounts[0]?.username})
                    </div>
                    <button onClick={handleLogout} style={{ padding: "8px 15px", backgroundColor: "#dc3545", color: "white", border: "none", borderRadius: "3px", cursor: "pointer" }}>
                        Cerrar Sesión
                    </button>
                </div>

                {/* SECCIÓN CREAR SOLICITUD */}
                <div style={{ border: "1px solid #ddd", padding: "15px", borderRadius: "5px", marginBottom: "20px" }}>
                    <h3>Crear Nueva Solicitud</h3>
                    <form onSubmit={crearSolicitud} style={{ display: "flex", gap: "10px" }}>
                        <input 
                            type="text" 
                            placeholder="Ej: Problema con el teclado" 
                            value={tipoSoporte} 
                            onChange={(e) => setTipoSoporte(e.target.value)} 
                            required 
                            style={{ flex: 1, padding: "8px" }}
                        />
                        <button type="submit" style={{ padding: "8px 15px", backgroundColor: "#0078D4", color: "white", border: "none", cursor: "pointer" }}>
                            Crear Ticket
                        </button>
                    </form>
                </div>

                {/* SECCIÓN LISTADO DE SOLICITUDES */}
                <div>
                    <h3>Listado de Solicitudes</h3>
                    <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
                        <thead>
                            <tr style={{ backgroundColor: "#f2f2f2" }}>
                                <th style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>ID</th>
                                <th style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>Usuario</th>
                                <th style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>Problema</th>
                                <th style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>Estado</th>
                                <th style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>Acciones (Roles)</th>
                            </tr>
                        </thead>
                        <tbody>
                            {solicitudes.length === 0 ? (
                                <tr><td colSpan="5" style={{ padding: "10px", textAlign: "center" }}>No hay solicitudes.</td></tr>
                            ) : (
                                solicitudes.map((ticket) => (
                                    <tr key={ticket.id}>
                                        <td style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>{ticket.id}</td>
                                        <td style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>{ticket.usuario}</td>
                                        <td style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>{ticket.tipoSoporte}</td>
                                        <td style={{ padding: "10px", borderBottom: "1px solid #ddd" }}><strong>{ticket.estado}</strong></td>
                                        <td style={{ padding: "10px", borderBottom: "1px solid #ddd" }}>
                                            {/* Renderizado condicional de botones según las reglas de negocio */}
                                            {ticket.estado === "CREADA" && (
                                                <button onClick={() => cambiarEstado(ticket.id, "ASIGNADA")}>Asignar</button>
                                            )}
                                            {ticket.estado === "ASIGNADA" && (
                                                <button onClick={() => cambiarEstado(ticket.id, "EN_PROCESO")}>Iniciar Proceso</button>
                                            )}
                                            {ticket.estado === "EN_PROCESO" && (
                                                <button onClick={() => cambiarEstado(ticket.id, "RESUELTA")}>Resolver</button>
                                            )}
                                            {ticket.estado === "RESUELTA" && (
                                                <button onClick={() => cambiarEstado(ticket.id, "CERRADA")}>Cerrar Ticket</button>
                                            )}
                                            {(ticket.estado === "CREADA" || ticket.estado === "ASIGNADA") && (
                                                <button onClick={() => cambiarEstado(ticket.id, "CANCELADA")} style={{ marginLeft: "5px", color: "red" }}>Cancelar</button>
                                            )}
                                        </td>
                                    </tr>
                                ))
                            )}
                        </tbody>
                    </table>
                </div>
            </AuthenticatedTemplate>
        </div>
    );
}

export default App;