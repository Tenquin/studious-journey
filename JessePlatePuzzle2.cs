using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class JessePlatePuzzle2 : MonoBehaviour 
{
	private bool oneHit;
	[SerializeField]
	private bool twoHit;
	private bool threeHit;
	private JessePlatePuzzle pp1;
	private JessePlatePuzzle3 pp3;

	// Use this for initialization
	void Start () 
	{
		gameObject.GetComponent<Renderer>().material.color = Color.red;
		pp1 = GameObject.Find("PressurePlate").GetComponent<JessePlatePuzzle>();
		pp3 = GameObject.Find("PressurePlate (2)").GetComponent<JessePlatePuzzle3>();
	}

	// Update is called once per frame
	void Update () 
	{
		oneHit = pp1.GetOne ();
		threeHit = pp3.GetThree ();
	}

	void OnTriggerEnter(Collider other)
	{
		if (other.tag == "Player") 
		{
			if (oneHit && twoHit && !threeHit) 
			{
				pp1.SetOneFalse ();
				pp3.SetThreeTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (oneHit && !twoHit && !threeHit) 
			{
				pp3.SetThreeTrue ();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (!oneHit && threeHit) 
			{
				pp3.SetThreeFalse();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (!oneHit && !threeHit) 
			{
				pp3.SetThreeTrue();
				gameObject.GetComponent<AudioSource> ().Play();
			}
			if (oneHit && threeHit) 
			{
				pp3.SetThreeFalse();
				gameObject.GetComponent<AudioSource> ().Play();
			}
		}
	}

	public void SetTwoFalse()
	{
		twoHit = false;
		gameObject.GetComponent<Renderer> ().material.color = Color.red;
	}
	public void SetTwoTrue()
	{
		twoHit = true;
		gameObject.GetComponent<Renderer> ().material.color = Color.green;
	}
	public bool GetTwo()
	{
		return twoHit;
	}
}
