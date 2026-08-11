@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():

    if request.method == "POST":

        password = request.form.get("password", "")

        if password == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect(url_for("admin"))

        return """
        <!DOCTYPE html>
        <html lang="az">
        <head>
            <meta charset="UTF-8">
            <title>Admin giriş</title>
            <style>
                body {
                    font-family: Arial;
                    background: #f3f6fb;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    margin: 0;
                }

                .box {
                    background: white;
                    padding: 40px;
                    border-radius: 20px;
                    width: 350px;
                    text-align: center;
                    box-shadow: 0 10px 30px rgba(0,0,0,.15);
                }

                input {
                    width: 100%;
                    box-sizing: border-box;
                    padding: 14px;
                    margin: 20px 0;
                    border-radius: 10px;
                    border: 1px solid #ddd;
                }

                button {
                    width: 100%;
                    padding: 14px;
                    background: #2563eb;
                    color: white;
                    border: none;
                    border-radius: 10px;
                    cursor: pointer;
                }

                .error {
                    color: red;
                    margin-bottom: 15px;
                }
            </style>
        </head>

        <body>

        <div class="box">

            <h1>🔐 Admin panel</h1>

            <p>Admin parolunu daxil edin</p>

            <div class="error">
                Parol yanlışdır.
            </div>

            <form method="POST">

                <input
                    type="password"
                    name="password"
                    placeholder="Admin parolu"
                    required
                >

                <button type="submit">
                    DAXİL OL
                </button>

            </form>

        </div>

        </body>
        </html>
        """

    return """
    <!DOCTYPE html>
    <html lang="az">
    <head>
        <meta charset="UTF-8">
        <title>Admin giriş</title>

        <style>
            body {
                font-family: Arial;
                background: #f3f6fb;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
            }

            .box {
                background: white;
                padding: 40px;
                border-radius: 20px;
                width: 350px;
                text-align: center;
                box-shadow: 0 10px 30px rgba(0,0,0,.15);
            }

            input {
                width: 100%;
                box-sizing: border-box;
                padding: 14px;
                margin: 20px 0;
                border-radius: 10px;
                border: 1px solid #ddd;
            }

            button {
                width: 100%;
                padding: 14px;
                background: #2563eb;
                color: white;
                border: none;
                border-radius: 10px;
                cursor: pointer;
            }
        </style>
    </head>

    <body>

    <div class="box">

        <h1>🔐 Admin panel</h1>

        <p>Nəticələrə baxmaq üçün daxil olun</p>

        <form method="POST">

            <input
                type="password"
                name="password"
                placeholder="Admin parolu"
                required
            >

            <button type="submit">
                DAXİL OL
            </button>

        </form>

    </div>

    </body>
    </html>
    """
