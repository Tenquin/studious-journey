using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class JessePlatePuzzle : MonoBehaviour 
{
	[SerializeField]
	private bool oneHit;
	private bool twoHit;
	private bool threeHit;
	private JessePlatePuzzle2 pp2;
	private JessePlatePuzzle3 pp3;

	// Use this for initialization
	void Start () 
	{
		gameObject.GetComponent<Renderer>().material.color = Color.red;
		pp2 = GameObject.Find("PressurePlate (1)").GetComponent<JessePlatePuzzle2>();
		pp3 = GameObject.Find("PressurePlate (2)").GetComponent<JessePlatePuzzle3>();
	}
	
	// Update is called once per frame
	void Update () 
	{
		twoHit = pp2.GetTwo ();
		threeHit = pp3.GetThree ();
	}

	void OnTriggerEnter(Collider other)
	{
		if (other.tag == "Player") 
		{
			if (twoHit) 
			{
				pp2.SetTwoFalse();
				oneHit = true;
				gameObject.GetComponent<Renderer> ().material.color = Color.green;
				gameObject.GetComponent<AudioSource> ().Play();
			}

			if (!twoHit) 
			{
				pp2.SetTwoTrue ();
				oneHit = false;
				gameObject.GetComponent<Renderer> ().material.color = Color.red;
				gameObject.GetComponent<AudioSource> ().Play();
			}
		}
	}
	
	public void SetOneFalse()
	{
		oneHit = false;
		gameObject.GetComponent<Renderer> ().material.color = Color.red;
	}
	public void SetOneTrue()
	{
		oneHit = true;
		gameObject.GetComponent<Renderer> ().material.color = Color.green;
	}
	public bool GetOne()
	{
		return oneHit;
	}
}
