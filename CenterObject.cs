using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.SceneManagement;

public class CenterObject : MonoBehaviour 
{

    /*	public static bool isOpen;
        public GameObject[] theSwitch;
        private static bool canSwitch;
        public GameObject switchText;
        public GameObject ClosedObject;
        public ControlBouncer controlBouncer;
        private PlayerStateMachine playerScript;
        public GameObject player;
        private Switch switchScript;
        [SerializeField]
        private static bool switchTwo;

        [SerializeField]
        private bool levelOne;
        [SerializeField]
        private bool levelTwo;
        [SerializeField]
        private bool levelJesse;
        [SerializeField]
        private bool levelDavid;
        [SerializeField]
        public bool level5;
        public Key key;
        // Use this for initialization
        void Start () 
        {
            if(levelTwo)
                switchScript = GameObject.FindGameObjectWithTag("Switch").GetComponent<Switch>();



            //bounceTrigger = GameObject.Find("Test Bouncers (2)").GetComponent<BounceTriggers> ();
        }

        // Update is called once per frame
        void Update () 
        {
    //		if(levelJesse)
    //			switchScript = GameObject.FindGameObjectWithTag("Switch").GetComponent<Switch>();
            if (player == null)
            {
                player = GameObject.Find("PlayerObject(Clone)");
                playerScript = player.gameObject.GetComponent<PlayerStateMachine>();
            }
            if (levelDavid )
            {
                key = GameObject.Find("Key").GetComponent<Key>();
            }
            if (levelTwo)
                switchTwo = switchScript.GetSwitchState ();

            /*if (levelJesse)
                switchTwo = switchScript.GetSwitchState ();

            if (isOpen) 
            {
                foreach(GameObject go in theSwitch) 
                {
                    go.SetActive(true);
                }
                ClosedObject.SetActive (false);
            }
            if (!isOpen) 
            {
                foreach (GameObject go in theSwitch) 
                {
                    go.SetActive (false);
                }
                ClosedObject.SetActive (true);
            }


            if (canSwitch && levelOne && Input.GetKeyDown (KeyCode.E)) 
            {
                isOpen = false;
                playerScript.GoTwo ();
                SceneManager.LoadScene ("VictoryScreen");
            }
            if (canSwitch && levelTwo && Input.GetKeyDown (KeyCode.E)) 
            {
                SceneManager.LoadScene ("VictoryScreen");
            }
            if (canSwitch && levelJesse && Input.GetKeyDown (KeyCode.E)) 
            {
                Destroy(player);
                SceneManager.LoadScene ("VictoryScreen");
                Cursor.visible = true;

            }
            if (canSwitch && levelDavid && Input.GetKeyDown(KeyCode.E))
            {
                isOpen = false;

                Destroy(player);
                SceneManager.LoadScene("VictoryScreen");
                Cursor.visible = true;
            }
            if (canSwitch && level5 && Input.GetKeyDown(KeyCode.E))
            {
                isOpen = false;

                Destroy(player);
                SceneManager.LoadScene("VictoryScreen");
                Cursor.visible = true;

            }
            //delete me
            if (Input.GetKeyDown (KeyCode.M)) 
            {
                isOpen = true;
            }
            if (levelOne && controlBouncer.GetComplete ()) 
            {
                isOpen = true;
            }
            if (levelDavid && key.Key1 == true)
            {
                isOpen = true;
            }
            if (level5 && controlBouncer.GetComplete())
            {
                isOpen = true;
            }
            if (levelTwo && switchTwo)
            {
                isOpen = true;
            }
            if (levelJesse && switchTwo)
            {
                isOpen = true;
            }
        }

        void OnTriggerEnter (Collider other)
        {
            if(other.name == "PlayerObject(Clone)" && isOpen)
            {
                canSwitch = true;
                switchText.SetActive (true);
            }
        }

        void OnTriggerExit ()
        {
            canSwitch = false;
            switchText.SetActive (false);
        }
        void Awake()
        {
            //theSwitch = GameObject.FindGameObjectsWithTag("theSwitch");
            if (!isOpen) 
            {
                foreach (GameObject go in theSwitch) 
                {
                    go.SetActive (false);
                }
                canSwitch = false;
                isOpen = false;
                switchText.SetActive (false);
            }

            //playerScript = GameObject.FindGameObjectWithTag ("Player").GetComponent<PlayerStateMachine> ();


            // Create a temporary reference to the current scene.
            Scene currentScene = SceneManager.GetActiveScene ();

            // Retrieve the name of this scene.
            string sceneName = currentScene.name;

            if (sceneName == "TestingGrounds") 
            {
                levelOne = true;
                levelTwo = false;
                canSwitch = false;
                levelDavid = false;
                level5 = false;
            } 

            else if (sceneName == "Level 2") 
            {
                levelOne = false;
                levelTwo = true;
                canSwitch = false;
                levelDavid = false;
                level5 = false;
            } 

            else if (sceneName == "Jesse's Level" || sceneName == "Jesse's Level Future") 
            {
                levelOne = false;
                levelTwo = false;
                levelJesse = true;
                canSwitch = false;
                levelDavid = false;
                level5 = false;
            }

            else if (sceneName == "DavidsLevel")
            {
                levelOne = false;
                levelTwo = false;
                levelJesse = false;
                levelDavid = true;
                level5 = false;
                canSwitch = false;
            }
            else if (sceneName == "Level 5")
            {
                levelOne = false;
                levelTwo = false;
                levelJesse = false;
                levelDavid = false;
                level5 = true;
                canSwitch = false;
            }
        }

        public void SetSwitch()
        {
            isOpen = true;
        }*/
}
