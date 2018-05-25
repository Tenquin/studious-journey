//References: http://answers.unity3d.com/questions/1117723/how-to-push-away-an-object-raycast-c.html
//References: http://answers.unity3d.com/questions/286821/transformtranslate-object-not-moving-toward-raycas.html
//References: http://answers.unity3d.com/questions/1192727/detect-if-raycast-hits-a-game-object.html

using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class AirTurret : MonoBehaviour 
{
	[SerializeField]
	//private GameObject thePlayer;

	// Use this for initialization
	void Start () 
	{
		//thePlayer = GameObject.FindGameObjectWithTag("Player");
	}
	
	// Update is called once per frame
	void Update () 
	{
		RaycastHit hit;
		float theDistance;


		Vector3 forward = transform.TransformDirection (Vector3.forward) * 10;
		//Debug.DrawRay (transform.position, forward, Color.blue, 10000);
		Debug.DrawLine (transform.position, forward, Color.blue, 10000);

		if (Physics.Raycast (transform.position, (forward), out hit))
		{
			if (hit.collider.gameObject.tag == "Player") 
			{
				theDistance = hit.distance;
				print (theDistance + " " + hit.collider.gameObject.name);
				//thePlayer.transform.Translate ((hit.point - transform.position) * Time.deltaTime);
				hit.transform.position += Vector3.forward * 100 * Time.deltaTime;
			}
		}
	}
}
