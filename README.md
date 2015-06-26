# Qu'est ce que c'est ?

Ceci est un programme qui g�n�re des fichiers Excel � partir d'un fichier SQL. Seul SQL Server est g�r�. Il peut m�me �tre envoy� par mail � un destinataire. Si votre requ�te remonte plus d'une grille de r�sultats, tous seront extraits.

# Comment �a marche ?

C'est simple, il y a un fichier de configuration `Config.xml` qui contient les informations n�cessaires � la cr�ation du fichier. Les param�tres sont �galement accessible via la ligne de commande `--param=valeur` pour pouvoir scripter par exemple l'extraction sur plusieurs bases de donn�es.

# Pourquoi ?

L'id�e de base �tait de g�n�rer un fichier Excel journalier pour suivre le statut d'un programme. Tous les jours, je devais r�g�n�rer le fichier, et l'envoyer � une personne. Bon en fait, une fois le programme d�velopp� (sur mes temps de pause et de repos), je me suis rendu compte que l'envoi de mail n'�tait pas possible, car la machine qui g�n�rait les rapports n'avait pas d'acc�s en SMTP. Du coup, j'ai une fonction que j'ai d�velopp�, sans jamais l'avoir test�e en production. Je l'ai test� en d�veloppement en revanche.

Pourquoi CarlogAG ExcelXmlWriter et pas les outils office pour g�n�rer un vrai fichier Excel ? Tout simplement pour avoir un minimum de d�pendances, et surtout, pas de Office install� sur la machine cible. Pratique pour un serveur.

# Comment ?

C'est � vous de voir... Une fois que vous avez param�tr� le programme via le fichier `Config.xml` et/ou que vous ayez d�fini tous les param�tres, il suffit de lancer le programme pour que le fichier Excel soit g�n�r�, et potentiellement envoy� par mail � la/aux personne(s) concern�e(s).

# Et les param�tres justement ?

De base, voici votre fichier `Config.xml` :

```
<?xml version="1.0" encoding="utf-8" ?>
<config>
  
</config>
```

C'est peu mais n�cessaire et suffisant pour �tre utilis� avec les param�tres de ligne de commande.

Chaque nouveau param�tre sera inscrit ainsi :
```
    <paramNom>Valeur</paramNom>
```

Voici la liste des param�tres g�r�s :

```
_dbserver_ : Serveur de base de donn�es avec instance. **Obligatoire**
_dblogin_ : Login � la base de donn�es. **Obligatoire si _dbtrusted_ est False / 0**
_dbpw_ : Mot de passe (malheureusement en clair) de la base de donn�es. **Obligatoire si _dbtrusted_ est False / 0**
_dbtrusted_ : Connexion en utilisant le login Windows (Trusted Auth) **Obligatoire** (Valeurs possibles : True / 1 / False / 0)
_dbdb_ : Base de donn�es par d�faut suite � la connexion. Rappelez vous qu'un USE peut la changer en cours de route !

_dateTimeFormat_ : Format des dates, car les dates sont retourn�es au format texte.

_mailsend_ : Doit-on envoyer le mail ?
_mailsmtp_ : Serveur SMTP d'envoi. **Obligatoire si mailsend est True / 1**
_mailsubject_ : Sujet du mail. **Obligatoire si mailsend est True / 1**
_mailbody_ : Contenu du mail. Il n'est pas param�trable. **Obligatoire si mailsend est True / 1**
_mailsender_ : Adresse envoyant le mail. **Obligatoire si mailsend est True / 1**
_mailrecipient_ : R�ceptionnaires du mail. Ils sont au format "Nom complet <adresse@example.com>" s�par�s par des ; et sans les " **Obligatoire si mailsend est True / 1**
_mailsmtpport_ : Port d'acc�s au serveur. **25 par d�faut**
_mailmustlogin_ : Doit-on se connecter au serveur de mail via un login/Mot de passe ?
_maillogin_ : Login de connexion au SMTP. **Obligatoire si mailmustlogin est True / 1**
_mailpw_ : Mot de passe de connexion au SMTP. **Obligatoire si mailmustlogin est True / 1**

_excelsheet_ : Doit-on extraire chaque r�sultat sur un nouvel onglet ? **Obligatoire**
_excelsqlfile_ : Fichier SQL � ex�cuter. **Obligatoire**
_excelfileprefix_ : Pr�fixe du nom du fichier Excel � extraire. Le fichier final sera de la forme _pr�fixe_yyyy_mm_dd.xml. Astuce : Le pr�fixe peut d�signer un chemin relatif ou complet. **Obligatoire**
_excellogprefix_ : Pr�fixe du log d'erreur. M�me r�gle qu'au dessus.
_excelsheetprefix_ : Pr�fixe du nom des onglets. Si excelsheet est False / 0, c'est le nom de l'onglet tel qu'il apparaitra. Sinon, les onglets seront nomm�s "Onglet 1", "Onglet 2"....
```

Rappellez vous qu'un param�tre sur la ligne de commande pr�vaut sur un param�tre du fichier `Config.xml`, que seul le premier param�tre du fichier Config.xml est pris en compte, alors que seul le dernier param�tre sur la ligne de commande est pris en compte.

Dernier d�tail, il y a un type de param�tres dynamique qui permet d'injecter des valeurs � des variables SQL Server. Ceux ci sont utilisables uniquement via la ligne de commande (si vous me trouvez l'int�r�t de pouvoir les param�trer via le fichier `Config.xml`, je ferais peut-�tre un effort...)

```
--param-type-_nom_=type
--param-value_nom_=valeur
```

_nom_ correspond au @variable dans votre fichier SQL (ne le pr�fixez pas par un @) et sera pass� en tant que param�tre � la commande SQL. Ne le d�clarez pas dans votre fichier SQL, le SqlCommand prenant en charge l'ex�cution de la requ�te s'en chargera pour vous.

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

Les types � taille fixe se transforment en type variable, donc n'h�sitez pas � rajouter des contr�les/conversions dans vos scripts. N'ayant pas de type date seule, ni temps seul en C#, tout est transform� en DateTime. Idem, pour le type image, qui n'est globalement qu'un type varbinary dans SQL.
