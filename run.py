from app import create_app, db
from app.models import User, Criteria, Subcriteria

app = create_app()

if __name__ == "__main__":
    # Initialize DB tables on first run
    with app.app_context():
        db.create_all()
        # Seed default criteria (if empty)
        if Criteria.query.count() == 0:
            defaults = [
                ("Nilai rata rata Raport", "benefit", 1),
                ("Nilai Sikap", "benefit", 2),
                ("Nilai Ketidakhadiran", "cost", 3),
                ("Nilai Estrakulikuler", "benefit", 4),
                ("Nilai Prestasi", "benefit", 5),
                ("Nilai Prakerin", "benefit", 6),
            ]
            for name, ctype, order in defaults:
                db.session.add(Criteria(name=name, ctype=ctype, display_order=order))
            db.session.commit()
    app.run(debug=True, host="0.0.0.0", port=5000)
