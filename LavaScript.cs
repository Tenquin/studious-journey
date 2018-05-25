using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class LavaScript : MonoBehaviour 
{
	public Transform playerSpawn;

    public GameObject player;
    public PlayerStateMachine stateMachine;

   // private FuelControl fuelRemaining;

    // Use this for initialization
    void Start () 
	{
      //  fuelRemaining = GameObject.Find("FuelMeter").GetComponent<FuelControl>();

    }
	
	// Update is called once per frame
	void Update () 
	{


        if (player == null)
        {
            player = GameObject.Find("PlayerObject(Clone)");
            stateMachine = player.GetComponent<PlayerStateMachine>();
        }

    }

	void OnTriggerEnter(Collider other)
	{
		//Debug.Log ("Trigger hit");

		if (other.tag == "Player") 
		{
            
               // stateMachine.SwitchToFloor();
               // stateMachine.OnFloor = true;
                stateMachine.zeroGravity = false;
                stateMachine.hasRespawned = true;
               // fuelRemaining.CurFuel = 0;

            

          
            other.transform.position = playerSpawn.transform.position;
		}
	}

    void OnTriggerExit(Collider other)
    {

        if (other.tag == "Player")
        {

            stateMachine.hasRespawned = false;

        }

    }
}
