from __future__ import annotations

from typing import Optional, cast
from flask import Blueprint, render_template, redirect, url_for, flash, request
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import login_user, logout_user, login_required, current_user
from sqlalchemy.exc import IntegrityError

from .models import User  # asumsi: model SQLAlchemy biasa
from .forms import LoginForm, RegisterForm
from . import db

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")


# ---------- Helpers untuk normalisasi input ----------
def normalize_email(value: Optional[str]) -> str:
    # Hilangkan spasi dan samakan huruf; fallback "" jika None
    return (value or "").strip().lower()


def normalize_name(value: Optional[str]) -> str:
    return (value or "").strip()


def normalize_password(value: Optional[str]) -> str:
    # Selalu kembalikan string (hashing & checking butuh str)
    return value or ""


# ---------- Routes ----------
@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    form = LoginForm()

    if form.validate_on_submit():
        email = normalize_email(form.email.data)
        password = normalize_password(form.password.data)

        if not email or not password:
            flash("Email dan kata sandi wajib diisi.", "danger")
            return render_template("login.html", form=form)

        # Cari user berdasarkan email yang sudah dinormalisasi
        user: Optional[User] = User.query.filter_by(email=email).first()

        # Pastikan password_hash ada sebelum check
        if user and getattr(user, "password_hash", None):
            if check_password_hash(user.password_hash, password):
                # Ambil remember_me jika ada di form; default False agar lebih aman
                remember_field = getattr(form, "remember_me", None)
                remember_flag = bool(getattr(remember_field, "data", False))
                login_user(user, remember=remember_flag)
                return redirect(url_for("main.dashboard"))

        flash("Email atau kata sandi salah.", "danger")

    return render_template("login.html", form=form)


@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))

    form = RegisterForm()

    if form.validate_on_submit():
        name = normalize_name(form.name.data)
        email = normalize_email(form.email.data)
        password = normalize_password(form.password.data)

        # Validasi sederhana tambahan (selain WTForms)
        if not name or not email or not password:
            flash("Nama, email, dan kata sandi wajib diisi.", "warning")
            return render_template("register.html", form=form)

        # Cek duplikasi email (case-insensitive) secara eksplisit
        if User.query.filter_by(email=email).first():
            flash("Email sudah terdaftar.", "warning")
            return redirect(url_for("auth.register"))

        try:
            # NOTE: Hindari keyword-args di konstruktor untuk memuaskan Pylance
            user = User()  # type: ignore[call-arg]
            # Set atribut satu per satu agar Pylance tidak protes
            user.name = name
            user.email = email
            user.password_hash = generate_password_hash(password)

            db.session.add(user)
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            # Jika ada unique constraint di DB, tangani di sini juga
            flash("Email sudah terdaftar.", "warning")
            return redirect(url_for("auth.register"))
        except Exception as e:
            db.session.rollback()
            flash("Terjadi kesalahan saat registrasi. Coba lagi.", "danger")
            # Opsional: log e
            return render_template("register.html", form=form)

        flash("Registrasi berhasil. Silakan login.", "success")
        return redirect(url_for("auth.login"))

    return render_template("register.html", form=form)


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Anda telah logout.", "info")
    return redirect(url_for("auth.login"))
