using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.SceneManagement;
public class JesseYouWin : MonoBehaviour 
{
	public GameObject SwitchText;
	public GameObject player;
	public GameObject playerSpawn;
	private bool switchHit;

	// Use this for initialization
	void Start () 
	{
		switchHit = false;
	}

	// Update is called once per frame
	void Update () 
	{
		if (player == null)
		{
			player = GameObject.Find("PlayerObject(Clone)");

		}
	}

	void OnTriggerStay (Collider other)
	{

		if (other.tag == "Player") {

			SwitchText.SetActive (true);

			if (Input.GetKeyDown (KeyCode.E)) 
			{
				other.transform.position = playerSpawn.transform.position;
				switchHit = true;
			}
		}
	}

	void OnTriggerExit(Collider other)
	{
		if (other.tag == "Player")
		{

			SwitchText.SetActive(false);
		}


	}

	public bool GetComplete()
	{
		return switchHit;
	}
}
