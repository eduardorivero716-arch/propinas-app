        table = pd.DataFrame(
            [
                {"Usuario": u, "Sucursal": r["name"], "Estado": "Activa" if r["active"] else "Desactivada", "Origen": r["source"]}
                for u, r in sorted(users.items())
            ]
        )
        st.dataframe(table, hide_index=True, width="stretch")

        # ---- Acciones
        st.subheader("Gestionar una sucursal")
        selected = st.selectbox(
            "Sucursal", options=sorted(users), format_func=lambda u: f"{users[u]['name']} ({u})", key="manage_select"
        )
        rec = users[selected]
        a, b = st.columns(2)

        with a:
            st.markdown(f"**Estado actual:** {'Activa' if rec['active'] else 'Desactivada'}")
            label = "Desactivar sucursal" if rec["active"] else "Activar sucursal"
            if st.button(label, key=f"toggle_{selected}"):
                err = set_active(selected, not rec["active"])
                if err:
                    st.error(err)
                else:
                    st.session_state["flash"] = (
                        f'Sucursal "{rec["name"]}" {"desactivada: perderá el acceso de inmediato" if rec["active"] else "activada"}.'
                    )
                    st.rerun()

        with b:
            with st.form(f"pw_form_{selected}", clear_on_submit=True):
                st.markdown("**Cambiar contraseña**")
                p1 = st.text_input("Nueva contraseña", type="password")
                p2 = st.text_input("Repetir contraseña", type="password")
                if st.form_submit_button("Cambiar contraseña"):
                    if p1 != p2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        err = set_password(selected, p1)
                        if err:
                            st.error(err)
                        else:
                            st.session_state["flash"] = f'Contraseña de "{rec["name"]}" actualizada.'
                            st.rerun()

        # ---- Porcentajes de reparto de la sucursal
        st.markdown("**Porcentajes de reparto de la sucursal**")
        cur = branch_pcts(rec)
        with st.form(f"pct_form_{selected}"):
            pc = st.columns(3)
            new_p = {}
            for col, (area, label) in zip(pc, PCT_AREAS):
                with col:
                    new_p[area] = st.number_input(f"{label} (%)", min_value=0.0, max_value=100.0, step=0.5,
                                                  value=float(cur[area]), format="%.2f", key=f"adm_pct_{selected}_{area}")
            if st.form_submit_button("Guardar porcentajes"):
                total = round(sum(new_p.values()), 6)
                if total != 100.0:
                    st.error(f"Los porcentajes suman {fmt_pct(total)}%. Deben sumar exactamente 100%.")
                else:
                    err = set_pcts(selected, new_p)
                    if err:
                        st.error(err)
                    else:
                        st.session_state["flash"] = f'Porcentajes de "{rec["name"]}" guardados.'
                        st.rerun()

        # ---- Aporte del bachero (Alfredo) de la sucursal
        st.markdown("**Aporte del bachero (Alfredo) de la sucursal**")
        cur_alf = branch_alfredo(rec)
        with st.form(f"alf_form_{selected}"):
            ac = st.columns(2)
            new_alf = {}
            for col, (area, label) in zip(ac, ALFREDO_AREAS):
                with col:
                    new_alf[area] = st.number_input(f"{label}: aporte por asistencia ($)", min_value=0, max_value=100000,
                                                    step=100, value=int(cur_alf[area]), key=f"adm_alf_{selected}_{area}")
            if st.form_submit_button("Guardar montos"):
                err = set_alfredo(selected, new_alf)
                if err:
                    st.error(err)
                else:
                    st.session_state["flash"] = f'Montos de Alfredo de "{rec["name"]}" guardados.'
                    st.rerun()

        # ---- Exportar para secrets
        st.subheader("Hacer los cambios permanentes")
        st.caption("Genera el bloque de configuración (con contraseñas ya cifradas) y pégalo en Settings > Secrets de tu app.")
        if st.button("Generar configuración para st.secrets"):
            toml = build_secrets_toml(users)
            st.code(toml, language="toml")
            st.download_button("Descargar como .toml", data=toml, file_name="sucursales_secrets.toml", mime="text/plain")


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================
def main():
    user = current_user()
    if not user:
        render_login()
        return

    with st.sidebar:
        st.markdown(f"**{user['name']}**")
        st.caption("Administrador" if user["role"] == "admin" else "Sucursal")
        page = "Cálculo de Propinas"
        if user["role"] == "admin":
            page = st.radio("Menú", ["Cálculo de Propinas", "Panel de Control de Acceso"], key="page")
        if st.button("Cerrar sesión"):
            logout()

    if page == "Panel de Control de Acceso" and user["role"] == "admin":  # doble verificación de rol
        render_admin_panel()
    else:
        render_tool(user)


main()
