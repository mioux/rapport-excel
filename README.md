# Qu'est ce que c'est ?

Ceci est un programme qui génère des fichiers Excel à partir d'un fichier SQL. Seul SQL Server est géré. Il peut même être envoyé par mail à un destinataire. Si votre requête remonte plus d'une grille de résultats, tous seront extraits.

# Comment ça marche ?

C'est simple, il y a un fichier de configuration `Config.xml` qui contient les informations nécessaires à la création du fichier. Les paramètres sont également accessible via la ligne de commande `--param=valeur` pour pouvoir scripter par exemple l'extraction sur plusieurs bases de données.

# Pourquoi ?

L'idée de base était de générer un fichier Excel journalier pour suivre le statut d'un programme. Tous les jours, je devais régénérer le fichier, et l'envoyer à une personne. Bon en fait, une fois le programme développé (sur mes temps de pause et de repos), je me suis rendu compte que l'envoi de mail n'était pas possible, car la machine qui générait les rapports n'avait pas d'accès en SMTP. Du coup, j'ai une fonction que j'ai développé, sans jamais l'avoir testée en production. Je l'ai testé en développement en revanche.

Pourquoi CarlogAG ExcelXmlWriter et pas les outils office pour générer un vrai fichier Excel ? Tout simplement pour avoir un minimum de dépendances, et surtout, pas de Office installé sur la machine cible. Pratique pour un serveur.

# Comment ?

C'est à vous de voir... Une fois que vous avez paramétré le programme via le fichier `Config.xml` et/ou que vous ayez défini tous les paramètres, il suffit de lancer le programme pour que le fichier Excel soit généré, et potentiellement envoyé par mail à la/aux personne(s) concernée(s).

# Et les paramètres justement ?

De base, voici votre fichier `Config.xml` :

```
<?xml version="1.0" encoding="utf-8" ?>
<config>
  
</config>
```

C'est peu mais nécessaire et suffisant pour être utilisé avec les paramètres de ligne de commande.

Chaque nouveau paramètre sera inscrit ainsi :
```
    <paramNom>Valeur</paramNom>
```

Voici la liste des paramètres gérés :

_dbserver_ : Serveur de base de données avec instance. **Obligatoire**
_dblogin_ : Login à la base de données. **Obligatoire si _dbtrusted_ est False / 0**
_dbpw_ : Mot de passe (malheureusement en clair) de la base de données. **Obligatoire si _dbtrusted_ est False / 0**
_dbtrusted_ : Connexion en utilisant le login Windows (Trusted Auth) **Obligatoire** (Valeurs possibles : True / 1 / False / 0)
_dbdb_ : Base de données par défaut suite à la connexion. Rappelez vous qu'un USE peut la changer en cours de route !

_dateTimeFormat_ : Format des dates, car les dates sont retournées au format texte.

_mailsend_ : Doit-on envoyer le mail ?
_mailsmtp_ : Serveur SMTP d'envoi. **Obligatoire si mailsend est True / 1**
_mailsubject_ : Sujet du mail. **Obligatoire si mailsend est True / 1**
_mailbody_ : Contenu du mail. Il n'est pas paramétrable. **Obligatoire si mailsend est True / 1**
_mailsender_ : Adresse envoyant le mail. **Obligatoire si mailsend est True / 1**
_mailrecipient_ : Réceptionnaires du mail. Ils sont au format "Nom complet <adresse@example.com>" séparés par des ; et sans les " **Obligatoire si mailsend est True / 1**
_mailsmtpport_ : Port d'accès au serveur. **25 par défaut**
_mailmustlogin_ : Doit-on se connecter au serveur de mail via un login/Mot de passe ?
_maillogin_ : Login de connexion au SMTP. **Obligatoire si mailmustlogin est True / 1**
_mailpw_ : Mot de passe de connexion au SMTP. **Obligatoire si mailmustlogin est True / 1**

_excelsheet_ : Doit-on extraire chaque résultat sur un nouvel onglet ? **Obligatoire**
_excelsqlfile_ : Fichier SQL à exécuter. **Obligatoire**
_excelfileprefix_ : Préfixe du nom du fichier Excel à extraire. Le fichier final sera de la forme _préfixe_yyyy_mm_dd.xml. Astuce : Le préfixe peut désigner un chemin relatif ou complet. **Obligatoire**
_excellogprefix_ : Préfixe du log d'erreur. Même règle qu'au dessus.
_excelsheetprefix_ : Préfixe du nom des onglets. Si excelsheet est False / 0, c'est le nom de l'onglet tel qu'il apparaitra. Sinon, les onglets seront nommés "Onglet 1", "Onglet 2"....

Rappellez vous qu'un paramètre sur la ligne de commande prévaut sur un paramètre du fichier `Config.xml`, que seul le premier paramètre du fichier Config.xml est pris en compte, alors que seul le dernier paramètre sur la ligne de commande est pris en compte.

Dernier détail, il y a un type de paramètres dynamique qui permet d'injecter des valeurs à des variables SQL Server. Ceux ci sont utilisables uniquement via la ligne de commande (si vous me trouvez l'intérêt de pouvoir les paramétrer via le fichier `Config.xml`, je ferais peut-être un effort...)

```
--param-type-_nom_=type
--param-value_nom_=valeur
```

_nom_ correspond au @variable dans votre fichier SQL (ne le préfixez pas par un @) et sera passé en tant que paramètre à la commande SQL. Ne le déclarez pas dans votre fichier SQL, le SqlCommand prenant en charge l'exécution de la requête s'en chargera pour vous.

type fait partie des types SQL Server, et sont convertis ainsi dans le programme :

```
tinyint => byte
smallint => Short
int => int
bigint => long
binary / varbinary / image => byte[]
bit => bool
datetime / smalldatetime / timestamp => DateTime
decimal / money / numeric / smallmoney => Decimal
float / real => double
varchar / char / nchar / nvarchar / text / ntext => string
uniqueidentifier => Guid
Tout le reste => object
```

Les types à taille fixe se transforment en type variable, donc n'hésitez pas à rajouter des contrôles/conversions dans vos scripts. N'ayant pas de type date seule, ni temps seul en C#, tout est transformé en DateTime. Idem, pour le type image, qui n'est globalement qu'un type varbinary dans SQL.
