using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class JessePlatePuzzle3 : MonoBehaviour 
{
	private bool oneHit;
	private bool twoHit;
	[SerializeField]
	private bool threeHit;
	private JessePlatePuzzle pp1;
	private JessePlatePuzzle2 pp2;

	// Use this for initialization
	void Start () 
	{
		gameObject.GetComponent<Renderer>().material.color = Color.red;
		pp1 = GameObject.Find("PressurePlate").GetComponent<JessePlatePuzzle>();
		pp2 = GameObject.Find("PressurePlate (1)").GetComponent<JessePlatePuzzle2>();
	}

	// Update is called once per frame
	void Update () 
	{
		oneHit = pp1.GetOne ();
		twoHit = pp2.GetTwo ();
	}

	void OnTriggerEnter(Collider other)
	{
		if (other.tag == "Player") 
		{
			if (oneHit && !twoHit && !threeHit) 
			{
				pp1.SetOneFalse();
				pp2.SetTwoTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (oneHit && !twoHit && threeHit)
			{
				pp1.SetOneTrue();
				pp2.SetTwoTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}	
			if (!oneHit && twoHit) 
			{
				pp1.SetOneTrue();
				pp2.SetTwoFalse();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (!oneHit && !twoHit && !threeHit) 
			{
				pp1.SetOneTrue ();
				pp2.SetTwoTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (!oneHit && !twoHit && threeHit) 
			{
				pp2.SetTwoTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (oneHit && twoHit) 
			{
				pp1.SetOneFalse();
				pp2.SetTwoFalse();
				gameObject.GetComponent<AudioSource> ().Play();
			}
		}
	}

	public void SetThreeFalse()
	{
		threeHit = false;
		gameObject.GetComponent<Renderer> ().material.color = Color.red;
	}
	public void SetThreeTrue()
	{
		threeHit = true;
		gameObject.GetComponent<Renderer> ().material.color = Color.green;
	}
	public bool GetThree()
	{
		return threeHit;
	}
}
