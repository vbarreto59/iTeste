    <style>
        header {
            background-color: #FFF; /* Cor de fundo */
            color: #000; /* Cor do texto */
            padding: 12px; /* Espaçamento interno */
            text-align: center; /* Alinhamento do texto */
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2); /* Sombra */
        }

        header h1 {
            margin: 0; /* Remove margem padrão do h1 */
            font-family: Arial, sans-serif; /* Fonte personalizada */
        }

        header nav {
            margin-top: 10px;
        }

        header nav a {
            color: white;
            text-decoration: none;
            margin: 0 10px;
        }

        header nav a:hover {
            text-decoration: underline;
        }
    </style>

       <header>
          <img src="img/logosunny.jpg" width="266" height="64" alt="SunnyImob">
          <h5>Usuário: <%=UCase(Session("Usuario"))%></h5>
       </header> 
       <br>