<?php
/**
 * API PHP ODT
 * PHP Library for ODT content manipulations
 * @author Florian Ajir <florian.ajir@adullact.org>
 * @created 15/11/13
 * @modified 11/12/13
 * @version 1.0 beta1
 */

class phpOdtApi
{

    /**
     * @var string chemin vers le fichier odt de référence
     */
    public $filepath;

    /**
     * @var string contenu au language XML du fichier content.xml
     */
    public $content;

    /**
     * @var character mode d'ouverture du fichier odt
     */
    public $mode;

    /**
     * @var DOMDocument arbre xml du document odt
     */
    private $dom;

    /**
     * @var ZipArchive instance de l'archive odt
     */
    private $zip;

    /**
     * @var DOMXpath
     */
    private $xpath;
    /**
     * @var array variables utilisateur utilisées dans le document
     */
    private $userFields;

    /**
     * @var array variables utilisateur déclarées dans le document
     */
    private $userFieldsDeclared;

    /**
     * @var array sections présentes dans le document
     */
    private $sections;

    /**
     * @var array conditions présentes dans le document
     */
    private $conditions;

    /**
     * @var array humanstyles présents dans le document
     */
    private $humanstyles;

    /**
     * @var array styles (auto) présents dans le document
     */
    private $autostyles;

    /**
     * @var array styles des nombres
     */
    private $numberstyles;

    /**
     * @var array variables déclarées dans le document
     */
    private $variablesDeclared;

    /**
     * @var array dictionnaire de correspondance des caractères spéciaux
     */
    private $dict = array(
        '_20_' => ' ',
        '_3d_' => '='
    );

    /**
     * filename of content xml in odt archive
     */
    const CONTENTXML = 'content.xml';

    /**
     * Initialisation
     * @param  string $file chemin du fichier sur le disque
     * @param  string $mode mode d'ouverture du fichier ('r' ou 'w')
     * @return bool succès de l'opération
     * @throws Exception si problème d'ouverture de l'odt
     */
    public function loadFromFile($file, $mode = 'r')
    {
        $this->filepath = $file;
        $this->mode = $mode;
        //Dézipage de l'odt
        $this->zip = new ZipArchive;
        if ($this->zip->open($this->filepath) === TRUE) {
            if ($this->content = $this->zip->getFromName(self::CONTENTXML)) {
                if ($mode == 'r')
                    $this->zip->close();
                if ($this->_loadDomFromXml())
                    return true;
            }
            $this->zip->close();
        }
        return false;
    }

    /**
     * Initialisation
     * @param  string $odt binaires du fichier odt
     * @param  string $mode mode d'ouverture du fichier ('r' ou 'w')
     * @param  bool $persistent garde le fichier sur le système
     * @return bool succès de l'opération
     * @throws Exception lors du chargement de l'odt
     */
    public function loadFromOdtBin(&$odt, $mode = 'r', $persistent = false)
    {
        $this->mode = $mode;
        //Dézipage de l'odt
        $this->zip = new ZipArchive();
        //Création de l'odt dans le dossier tmp
        $this->filepath = tempnam("/tmp", "WD_ODT_");
        file_put_contents($this->filepath, $odt);
        if ($this->zip->open($this->filepath) === TRUE) {
            if ($this->content = $this->zip->getFromName(self::CONTENTXML)) {
                if (!$this->_loadDomFromXml())
                    return false;
                if ($mode == 'r')
                    $this->zip->close();
            } else {
                $this->zip->close();
                return false;
            }
        } else {
            return false;
        }
        if (!$persistent) {
            //Suppression du fichier odt temporaire créé
            unlink($this->filepath);
            $this->filepath = null;
        }
        return true;
    }

    /**
     * Initialise l'attribut $dom à partir du contenu XML de $content
     */
    private function _loadDomFromXml()
    {
        if ($this->dom = DOMDocument::loadXML($this->content)) {
            $this->dom->formatOutput = true;
            self::_reloadDOMXpath();
            return true;
        }
        return false;
    }

    /**
     * Initialise l'attribut $dom à partir du contenu XML de $content
     */
    private function _saveXmlFromDom()
    {
        $this->content = $this->dom->saveXML();
        return !empty($this->content);
    }

    /**
     * Initialise l'attribut $dom à partir du contenu XML de $content
     */
    private function _reloadDOMXpath()
    {
        $this->xpath = new DOMXpath($this->dom);
    }

    /**
     * Sauvegarde les modifications effectuées sur l'arbre xml content du document odt
     * @param string $filepath chemin du nouveau fichier odt (si null, écrase l'original)
     * @param bool $return envoi l'odt en retour de la fonction
     * @return bool|string
     */
    public function save($return = false, $filepath = null)
    {
        self::_apply();
        if ($this->mode != 'w')
            return false;
        if ($this->_saveXmlFromDom()) {
            $this->zip->deleteName(self::CONTENTXML);
            $this->zip->addFromString(self::CONTENTXML, $this->content);
            $this->zip->close();
            if (!empty($filepath)) {
                if (file_exists($this->filepath)) {
                    //Copie le document vers la nouvelle destination
                    copy($this->filepath, $filepath); // Attention, Si le fichier de destination existe déjà, il sera écrasé.
                    //unlink($this->filepath); //Suppression du premier fichier ?
                    $this->filepath = $filepath;
                } else {
                    return false;
                }
            }
            if ($return)
                return file_get_contents($this->filepath);
            else
                return true;
        }
        return false;
    }

    /**
     * @return string numéro de version du document
     */
    public function getVersion(){
        $docNode = $this->xpath->query('/office:document-content')->item(0);
        return $docNode->getAttribute('office:version');
    }

    /**
     * Met fin à l'instance
     */
    public function close()
    {
        $this->zip->close();
    }

    /**
     * Récupère sous forme d'un tableau toutes les valeurs des attributs
     * @param string $tag nom des tags dans lesquels chercher les valeurs des attributs
     * @param string $attr nom de l'attribut à chercher
     * @return array liste d'attributs
     */
    private function _getAttributes($tag, $attr)
    {
        $elems = array();
        foreach ($this->dom->getElementsByTagNameNS('*', $tag) as $element) {
            if ($element->hasAttribute($attr))
                $elems[] = $element->getAttribute($attr);
        }
        return $elems;
    }

    /**
     * @param string $string to decode
     * @return string decoded
     */
    private function _decodeEntities($string)
    {
        return str_replace(array_keys($this->dict), array_values($this->dict), $string);
    }

    /**
     * Récupère les noms des sections du document
     * @return array
     */
    public function getSections()
    {
        $this->sections = $this->_getAttributes('section', 'text:name');
        return $this->sections;
    }

    /**
     * Récupère sous forme d'un tableau les noeuds cherchés
     * @param string $section section à chercher
     * @return array liste de noeuds
     */
    private function _getSectionNodes($section)
    {
        $elems = array();
        foreach ($this->dom->getElementsByTagNameNS('*', 'section') as $node) {
            if ($node->hasAttribute('text:name') && $node->getAttribute('text:name') == $section)
                $elems[] = $node;
        }
        return $elems;
    }

    /**
     * Le fichier contient-il la section désignée par son nom en argument
     * @param string $section nom de la variable
     * @return bool section true si présente dans le document, false dans le cas contraire
     */
    public function hasSection($section)
    {
        if (empty($this->sections))
            $this->sections = $this->getSections();
        return in_array($section, $this->sections);
    }

    /**
     * Compte le nombre d'occurrences de la section dans le document
     * @param $section nom de la section
     * @return integer nombre d'occurrences
     */
    public function countSection($section)
    {
        $count = 0;
        $duplicates = array_count_values($this->sections);

        if (!empty($duplicates[$section]))
            $count = $duplicates[$section];
        elseif ($this->hasSection($section))
            $count = 1;

        return $count;
    }

    /**
     * TODO
     * Retourne la section contenant la section en paramètre
     * @throws Exception si plusieurs parents (section non unique)
     * @param string $section section pour laquelle trouver le parent
     * @return string section parent
     */
    public function getParentSection($section)
    {
        $node = null;
        $nb = 0;
        $parent = 'Document';
        foreach ($this->dom->getElementsByTagNameNS('*', 'section') as $element) {
            if ($element->hasAttribute('text-name') && $element->getAttribute('text-name') == $section) {
//                $node = $element;
                $nb++;
            }
        }
        if ($nb > 1) throw new Exception("Plusieurs sections identiques ont été rencontrées dans le document");
        if ($nb != 1) return $parent;
        //TODO: Trouver avec xpath le noeud parent
//        $xPath = $node.getNodePath();
        return $parent;
    }


    /**
     * Récupère les noms des variables utilisateur du document
     * @return array
     */
    public function getUserFields()
    {
        $this->userFields = $this->_getAttributes('user-field-get', 'text:name');
        return $this->userFields;
    }


    /**
     * Récupère sous forme d'un tableau les noeuds cherchés
     * @param string $userField userField à chercher
     * @return array liste de noeuds
     */
    private function _getUserFieldNodes($userField)
    {
        $elems = array();
        foreach ($this->dom->getElementsByTagNameNS('*', 'user-field-get') as $node) {
            if ($node->hasAttribute('text:name') && $node->getAttribute('text:name') == $userField)
                $elems[] = $node;
        }
        return $elems;
    }

    /**
     * Le fichier contient-il la variable utilisateur désignée par son nom en argument
     * @param string $userField nom de la variable
     * @return bool variable true si présente dans le document, false dans le cas contraire
     */
    public function hasUserField($userField)
    {
        if (empty($this->userFields))
            $this->userFields = $this->getUserFields();
        return in_array($userField, $this->userFields);
    }

    /**
     * Le fichier contient-il au moins une des variables utilisateur désignées par son nom en argument
     * @param array|parameters_liste $userFields liste des noms des variables
     * @return bool true si une des variables est présente dans le document, false dans le cas contraire
     */
    public function hasUserFields($userFields)
    {
        if (!is_array($userFields)) $userFields = func_get_args();
        if (empty($this->userFields))
            $this->userFields = $this->getUserFields();
        foreach($userFields as $userField)
            if (in_array($userField, $this->userFields)) return true;
        return false;
    }

    /**
     * Compte le nombre d'occurrences de la variable de type userField dans le document
     * @param string $userField nom de la variable
     * @return integer nombre d'occurrences
     */
    public function countUserFields($userField)
    {
        $count = 0;
        $duplicates = array_count_values($this->userFields);
        if (!empty($duplicates[$userField]))
            $count = $duplicates[$userField];
        elseif ($this->hasUserField($userField))
            $count = 1;

        return $count;
    }

    /**
     * Retourne la ou les sections dans lesquelles se trouvent la variable
     * @param string $userField
     * @return array
     */
    public function getUserFieldSectionContainer($userField)
    {
        $sections = array();
        $nodes = $this->_getUserFieldNodes($userField);
        foreach ($nodes as $node) {
            $parent = $node->parentNode;
            while (!($parent->nodeName == 'text:section' && $parent->hasAttribute('text:name'))) {
                $parent = $parent->parentNode;
                if ($parent->nodeName == 'office:body') break;
            }
            if ($parent->nodeName != 'office:body')
                $sections[] = $parent->getAttribute('text:name');
        }
        return $sections;
    }

    /**
     * Compte le nombre d'occurrences de la variable de type userField dans la section désignée
     * @param string $userField nom de la variable
     * @param string $section nom de la section
     * @return integer nombre d'occurrences
     */
    public function countUserFieldsInSection($userField, $section)
    {
        $count = 0;
        $blocs = $this->_getSectionNodes($section);
        foreach ($blocs as $bloc) {
            foreach ($bloc->getElementsByTagNameNS('*', 'user-field-get') as $node) {
                if ($node->hasAttribute('text:name') && $node->getAttribute('text:name') == $userField)
                    $count++;
            }
        }
        return $count;
    }

    /**
     * Vérifie la présence d'une variable dans une section
     * @param string $userField la variable à rechercher
     * @param string $section le conteneur à rechercher
     * @return bool la variable existe t'elle dans cette section
     */
    public function hasUserFieldsInSection($userField, $section)
    {
        return ($this->countUserFieldsInSection($userField, $section) > 0);
    }

    /**
     * Récupère les noms des variables utilisateur déclarées dans le document (utilisées ou non)
     * @return array
     */
    public function getUserFieldsDeclared()
    {
        $this->userFieldsDeclared = array();
        $decls = $this->xpath->query('/office:document-content/office:body/office:text/text:user-field-decls/text:user-field-decl');
        foreach ($decls as $decl) {
            $this->userFieldsDeclared[$decl->getAttribute('text:name')] = array();
            foreach ($decl->attributes as $attr => $domattr) {
                $this->userFieldsDeclared[$decl->getAttribute('text:name')][$attr] = $domattr->value;
            }
        }
        return $this->userFieldsDeclared;
    }

    /**
     * la variable utilisateur est elle déclarée dans le document ?
     * @param string $name nom de la variable utilisateur
     * @return bool
     */
    public function hasUserFieldDeclared($name)
    {
        return array_key_exists($name, self::getUserFieldsDeclared());
    }

    /**
     * Récupère les noms des variables utilisateur déclarées dans le document mais non utilisées
     * Peut être les "_separator" pour les itérations sur les listes
     * @return array
     */
    public function getUserFieldsNotUsed()
    {
        $diff = array();
        if (empty($this->userFieldsDeclared))
            $this->userFieldsDeclared = $this->getUserFieldsDeclared();
        if (empty($this->userFields))
            $this->userFields = $this->getUserFields();
        foreach ($this->userFieldsDeclared as $name => $attrs)
            if (!in_array($name, $this->userFields))
                $diff[] = $name;

        return $diff;
    }

    /**
     * Récupère les noms des variables déclarées
     * Attention : celles-ci ne sont pas des variables "user-field"!!
     * @return array
     */
    public function getVariablesDeclared()
    {
        $this->variablesDeclared = $this->_getAttributes('variable-decl', 'text:name');
        return $this->variablesDeclared;
    }

    /**
     * Récupère les noms des styles du document (noms donnés par l'auteur du document)
     * @return array
     */
    public function getHumanStyles()
    {
        $this->humanstyles = array();
        $allStyles = $this->_getAttributes('style', 'style:parent-style-name');
        foreach ($allStyles as $style) {
            //Caractères spéciaux
            $style = $this->_decodeEntities($style);
            //Si pas encore dans le tableau
            if (!in_array($style, $this->humanstyles))
                $this->humanstyles[] = $style;
        }
        return $this->humanstyles;
    }

    /**
     * Le fichier contient-il le style désignée en argument
     * @param string $style
     * @return bool true si présent dans le document, false dans le cas contraire
     */
    public function hasHumanStyle($style)
    {
        if (empty($this->humanstyles))
            $this->humanstyles = $this->getHumanStyles();
        return in_array($style, $this->humanstyles);
    }

    /**
     * Récupère les conditions présentes dans le document
     * @return array
     */
    public function getConditions()
    {
        $this->conditions = str_replace('ooow:', '', $this->_getAttributes('*', 'text:condition'));
        return $this->conditions;
    }

    /**
     * Le fichier contient-il la condition désignée en argument
     * @param string $condition nom de la condition
     * @return bool true si présente dans le document, false dans le cas contraire
     */
    public function hasCondition($condition)
    {
        if (empty($this->conditions))
            $this->conditions = $this->getConditions();
        return in_array($condition, $this->conditions);
    }

    /**
     * Le fichier contient-il le noeud de déclaration des variables utilisateur
     * @return bool section true si présente dans le document, false dans le cas contraire
     */
    private function _hasUserFieldsDeclarationNode()
    {
        $declarations = $this->xpath->query("/office:document-content/office:body/office:text/text:user-field-decls");
        return ($declarations->length != 0);
    }

    /**
     * @param string $name nom de la variable utilisateur
     * @param string $type type de donnée {string, float, date}
     * @return bool succès
     */
    public function declareUserField($name, $type = 'string')
    {
        if (!self::_hasUserFieldsDeclarationNode()) {
            $declarationsElement = $this->dom->createElement('text:user-field-decls');
            //Trouver le noeud office:text
            $officeText = $this->xpath->query("/office:document-content/office:body/office:text")->item(0);
            if (!$officeText) return false;
            //Référence pour le placement du noeud de déclaration
            $ref = $this->xpath->query("/office:document-content/office:body/office:text/text:sequence-decls")->item(0)->nextSibling;
            if (!$ref) return false;
            //Ajouter le nouveau noeud de déclarations à l'endroit adéquat
            $declarations = $officeText->insertBefore($declarationsElement, $ref);
        } else {
            $declarations = $this->xpath->query("/office:document-content/office:body/office:text/text:user-field-decls")->item(0);
        }
        if (!$declarations) return false;
        $userFieldDecl = $this->dom->createElement('text:user-field-decl');
        $userFieldDecl = $declarations->appendChild($userFieldDecl);
        $userFieldDecl->setAttribute('office:value-type', $type);
        $value = ($type == 'string') ? $name : '0';

        if ($type == 'float')
            $userFieldDecl->setAttribute('office:value', $value);
        else
            $userFieldDecl->setAttribute("office:$type-value", $value);

        $userFieldDecl->setAttribute('text:name', $name);
        return true;
    }

    /**
     * @param string $type
     * @param string $name
     * @param string $parentname
     * @return bool
     */
    public function addStyle($type = 'section', $name = '', $parentname = '')
    {
        //Trouver le noeud style:style
        $parent = $this->xpath->query("/office:document-content/office:automatic-styles")->item(0);

        $style = $this->dom->createElement('style:style');
        $style = $parent->appendChild($style);
        if (empty($name)) {
            switch ($type){
                case 'section':
                    $key = 'Sect';
                    break;
                case 'graphic':
                    $key = 'fr';
                    break;
                case 'paragraph':
                    $key = 'P';
                    break;
                case 'text':
                    $key = 'T';
                    break;
                default:
                    $key = strtoupper($type[0]);
            }
            $styles = self::getAutoStyles();
            for ($i = 1; $i < count($styles[$type]); $i++) {
                if (!in_array($key.$i, $styles[$type]))
                    break;
            }
            $name = $key.$i;
        }
        $style->setAttribute('style:name', $name);
        $style->setAttribute('style:family', $type);
        if (!empty($parentname))
            $style->setAttribute('style:parent-style-name', $parentname);
//        self::_apply();
        return $name;
    }

    /**
     * Add a userfield to the end of the document, can be insert in an existing or a new section
     * @param string $userfield name (and value if string type and not already declared) of the userfield
     * @param string $type type of the userfield
     * @param string $section name of the section that must contain the userfield
     */
    public function appendUserField($userfield, $type = 'string', $section = null)
    {
        //Test si la variable est déclarée
        if (!self::hasUserFieldDeclared($userfield)) {
            //Sinon la déclarer
            self::declareUserField($userfield, $type);
            $value = $userfield;
        }else{
            $declared = self::getUserFieldsDeclared();
            if ($type != $declared[$userfield]['value-type'])
                $type = $declared[$userfield]['value-type'];
            $value = ($type == 'float') ? $declared[$userfield]['value'] : $declared[$userfield][$type.'-value'];
        }
        if (!empty($section)) {
            if (!self::hasSection($section)) {
                $parent = self::appendSection($section);
            } else {
                $sections = self::_getSectionNodes($section);
                $parent = $sections[0];
            }
            $parent = $parent->lastChild;
        } else {
            $parent = $this->xpath->query("/office:document-content/office:body/office:text")->item(0);
        }
        $uf = $this->dom->createElement('text:user-field-get');
        $uf = $parent->appendChild($uf);
        $uf->nodeValue = $value;
        if ($type == 'float') {
            $styles = self::getNumberStyles();
            if (!empty($styles))
                $uf->setAttribute('style:data-style-name', $styles[0]); //style : Standard
        }
//        else {
//            $styles = self::getAutoStyles();
//            $uf->setAttribute('style:data-style-name', $styles['paragraph'][0]); //style : Standard
//        }
        $uf->setAttribute('text:name', $userfield);
        self::_apply();
    }

    /**
     * Ajoute une nouvelle section à la fin du document (et déclare un style de section pour celle ci)
     * @param $sectionname
     * @return DOMElement|DOMNode
     */
    public function appendSection($sectionname)
    {
        $body = $this->xpath->query("/office:document-content/office:body/office:text")->item(0);

        $section = $this->dom->createElement('text:section');
        $section = $body->appendChild($section);
        $style = self::addStyle();
        $section->setAttribute('text:style-name', $style);
        $section->setAttribute('text:name', $sectionname);
        $p = $this->dom->createElement('text:p');
        $p = $section->appendChild($p);
        // FIXME : remplacer par $stylep = self::getHumanStyles(); ?
        $styles = self::getAutoStyles();
        $p->setAttribute('text:style-name', $styles['paragraph'][0]); //style : Standard

//        self::_apply();
        return $section;
    }

    /**
     * Applique les modifications du dom
     */
    private function _apply()
    {
        self::_saveXmlFromDom();
        self::_loadDomFromXml();
    }

    /**
     * Récupère les noms des styles du document
     * @return array
     */
    public function getAutoStyles()
    {
        $this->autostyles = array(
            'paragraph' => array(),
            'section' => array(),
            'text' => array(),
            'graphic' => array(),
        );
        $styles = $this->xpath->query('/office:document-content/office:automatic-styles/style:style');
        foreach ($styles as $style) {
            if ($style->hasAttribute('style:family')) {
                //Si pas encore dans le tableau
                if (empty($this->autostyles[$style->getAttribute('style:family')])
                    || !in_array($style->getAttribute('style:name'), $this->autostyles[$style->getAttribute('style:family')])
                )
                    $this->autostyles[$style->getAttribute('style:family')][] = $style->getAttribute('style:name');
            }
        }
        return $this->autostyles;
    }

    /**
     * Récupère les noms des styles du document
     * @return array
     */
    public function getNumberStyles()
    {
        $this->numberstyles = array();
        $styles = $this->xpath->query('/office:document-content/office:automatic-styles/number:number-style');
        foreach ($styles as $style) {
            //Si pas encore dans le tableau
            if (empty($this->autostyles) || !in_array($style->getAttribute('style:name'), $this->numberstyles))
                $this->numberstyles[] = $style->getAttribute('style:name');
        }
        return $this->numberstyles;
    }

    /**
     * Le fichier contient-il le style désignée en argument
     * @param string $style
     * @return bool true si présent dans le document, false dans le cas contraire
     */
    public function hasAutoStyle($style)
    {
        if (empty($this->autostyles))
            $this->autostyles = $this->getAutoStyles();
        return in_array($style, $this->autostyles);
    }

}
