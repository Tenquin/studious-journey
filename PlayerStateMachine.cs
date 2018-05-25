//references:http://answers.unity3d.com/questions/1173303/how-to-check-which-scene-is-loaded-and-write-if-co.html
//references:http://answers.unity3d.com/questions/982403/how-to-not-duplicate-game-objects-on-dontdestroyon.html
//references:http://answers.unity3d.com/questions/1230216/a-proper-way-to-pause-a-game.html

using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;
using UnityEngine.SceneManagement;

public class PlayerStateMachine : MonoBehaviour
{


    [SerializeField]
    private FuelControl UseFuel;
    [SerializeField]
    private float DecreasedAmount = 0.2f;
    [SerializeField]
    private float IncreasedAmount = 0.2f;
    [SerializeField]
    private bool isJumping = false;
    [SerializeField]
    private float jumpTimer = 1.25f;
    //    private bool atTop;
    public bool canClimb;
    [SerializeField]
    private float startSpeed;
    [SerializeField]
    private float sprintSpeed;
    [SerializeField]
    private float moveSpeed;
    [SerializeField]
    private float rotateSpeed;
    private CharacterController playerController;
    // Adding Camera Game Objects
    public GameObject FPSCam;
    //Adding in better aiming
    private float CurX = 0f;
    // the current value of y used to manipulate camera movement.
    private float CurY = 0f;
    // Min height the camera can go
    private const float Y_ANGLE_MIN = 1.0f;
    // Max height the camera can go
    private const float Y_ANGLE_MAX = 100.0f;
    private const float X_ANGLE_MIN = 01.0f;
    // Max height the camera can go
    private const float X_ANGLE_MAX = 50.0f;
    private float XSens = 0.5f;
    public bool OnFloor;
    public bool isFuture;
    [SerializeField]
    public bool isOnPlatform;
    // [SerializeField]
    public bool hasPressed = false;
    private static PlayerStateMachine playerInstance;
   // private GameObject[] pauseMenu;
    //[SerializeField]
    //public GameObject pauseMenuScreen;
    public bool isPaused;
    public bool zeroGravity;
    public bool hasFuel;
    public float speed = 6.0F;
    public float jumpSpeed = 8.0F;
    public float gravity;
    public float gravityValue;
    private Vector3 moveDirection = Vector3.zero;
    public bool onRoof;
    public bool onground;
    public bool inair;
    public bool flipped;
    public bool notflipped;
    public bool canJump;
    public bool hasRespawned;
    public bool inPast;
    public Transform from;
    public Transform to;
    public float lspeed = 0.1F;
    private AudioSource footSteps;
    public int orbs;
    public int Crystals;
    public int Souls;
	public bool pickedUpFuel;
    //Create an enum to hold all possible states
    enum PlayerStates
    {
        jetPack,
        idle,
        basicmovement,
        climbing,
        Inair,
        Num_States
    }
    //set initial state
    [SerializeField]
    private PlayerStates curState = PlayerStates.idle;
    //create a dictionary for states to be added to so the can be called. 
    private Dictionary<PlayerStates, Action> ps = new Dictionary<PlayerStates, Action>();
    void Start()
    {
        footSteps = gameObject.GetComponent<AudioSource>();
        inPast = false;
        FPSCam = GameObject.Find("FPCamera");
        OnFloor = true;
        onRoof = false;
        inair = false;
        Cursor.lockState = CursorLockMode.Confined;
        ps.Add(PlayerStates.jetPack, new Action(StateJetPack));
        ps.Add(PlayerStates.idle, new Action(Stateidle));
        ps.Add(PlayerStates.basicmovement, new Action(StateBasicMovement));
        ps.Add(PlayerStates.climbing, new Action(StateClimbing));
        ps.Add(PlayerStates.Inair, new Action(StateInair));
		pickedUpFuel = false;
    }

    void Update()
    {
		

        if (Input.GetKeyDown(KeyCode.F))
        {
            UseFuel.CurFuel -= DecreasedAmount;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount;
            inPast = !inPast;
            
        }
        if (flipped && !zeroGravity)
        {
            gravity = -gravityValue;
        }

        if (!flipped && !zeroGravity)
        {
            gravity = gravityValue;
        }
        if (!isPaused)
        {
            transform.Rotate(0, Input.GetAxis("Mouse X") * rotateSpeed, 0);
            // gets tge Y axus from the mouse for the rotaion so the camera can look up or down 
            CurX += Input.GetAxis("Mouse Y") * -rotateSpeed;
            // locks the camera
            CurX = Mathf.Clamp(CurX, -60f, 60);
            // rotates the camear up or down
            FPSCam.transform.localRotation = Quaternion.Euler(CurX, 0, 0);
        }
        /*if (!zeroGravity)
        {
            gravity = 30;
        }*/
        CharacterController controller = GetComponent<CharacterController>();


        if (controller.isGrounded)
        {

            canJump = true;
            inair = false;

            moveDirection = new Vector3(Input.GetAxis("Horizontal"), 0, Input.GetAxis("Vertical"));
            moveDirection = transform.TransformDirection(moveDirection);
            moveDirection *= moveSpeed;

            // moveDirection.y -= gravity * Time.deltaTime;

        }

        //needs to not be done when grounded anymore
        controller.Move(moveDirection * Time.deltaTime);

        if (onRoof && !zeroGravity && curState != PlayerStates.jetPack && flipped && canJump)
        {
            if (Input.GetKey(KeyCode.Space))
            {

                moveDirection.y = -jumpSpeed;
                onground = false;
                curState = PlayerStates.Inair;

            }

        }

        if (onground && !zeroGravity && curState != PlayerStates.jetPack && !flipped && canJump)
        {
            if (Input.GetKey(KeyCode.Space))
            {
                moveDirection.y = jumpSpeed;
                onground = false;
                curState = PlayerStates.Inair;
            }
        }
        if (inair)
        {
            moveDirection.y -= gravity * Time.deltaTime;
            canJump = false;
            if (controller.isGrounded)
            {
                curState = PlayerStates.idle;
                canJump = true;
            }
        }


        // moveDirection.y -= gravity * Time.deltaTime;
        controller.Move(moveDirection * Time.deltaTime);
        if (flipped)
        {
            if (curState != PlayerStates.jetPack)
            {
                moveDirection.y -= gravity * Time.deltaTime;
            }
        }
        if (!flipped)
        {
            if (curState != PlayerStates.jetPack)
            {
                moveDirection.y -= gravity * Time.deltaTime;
            }
        }

        if (Input.GetKeyDown(KeyCode.G) && inPast == true)
        {
            //Debug.Log("can");
            zeroGravity = !zeroGravity;
        }
       /* if (transform.Find("PauseMenu"))
        {
            //Debug.Log("Found it");
        }
        if (pauseMenuScreen == null)
        {
            pauseMenuScreen = GameObject.Find("PauseMenu");
        }
        */
        if (isJumping == true)
        {
            jumpTimer -= Time.deltaTime;
        }
        if (isJumping == false)
        {
            jumpTimer = 1.5f;
        }
        ps[curState].Invoke();
        if (UseFuel == null)
        {
            UseFuel = GameObject.Find("FuelMeter").GetComponent<FuelControl>();
        }
        if (FPSCam == null)
        {

            FPSCam = GameObject.Find("FPCamera");
        }

        if (Input.GetKeyDown(KeyCode.LeftShift))
        {
            moveSpeed = sprintSpeed;

        }
        if (Input.GetKeyUp(KeyCode.LeftShift))
        {
            moveSpeed = startSpeed;

        }
    }
    //ControlIdle State
    public void Stateidle()
    {
        onground = true;
        inair = false;

        if (zeroGravity == true)
        {
            curState = PlayerStates.jetPack;
        }
        isOnPlatform = false;
        //Debug.Log ("I am Idle");
        if (Input.GetKeyDown(KeyCode.W))
        {
            //footSteps.Play ();
            curState = PlayerStates.basicmovement;
        }
        if (Input.GetKeyDown(KeyCode.S))
        {

            curState = PlayerStates.basicmovement;
        }
        if (Input.GetKeyDown(KeyCode.A))
        {

            curState = PlayerStates.basicmovement;
        }
        if (Input.GetKeyDown(KeyCode.D))
        {

            curState = PlayerStates.basicmovement;
        }

        if (canClimb == true && Input.GetKeyDown(KeyCode.E))
        {
            //climbText.SetActive (true);
            curState = PlayerStates.climbing;
        }
    }
    void StateBasicMovement()
    {
        if (Input.GetKey(KeyCode.W) && footSteps.isPlaying == false)
        {
            footSteps.Play();
        }
        if (Input.GetKey(KeyCode.A) && footSteps.isPlaying == false)
        {
            footSteps.Play();
        }
        if (Input.GetKey(KeyCode.S) && footSteps.isPlaying == false)
        {
            footSteps.Play();
        }
        if (Input.GetKey(KeyCode.D) && footSteps.isPlaying == false)
        {
            footSteps.Play();
        }
        //Debug.Log (playerController.isGrounded);
        if (zeroGravity == true)
        {
            curState = PlayerStates.jetPack;
        }
        // check to see if the isOnPlatform is Checked
        if (isOnPlatform == true)
        {
            curState = PlayerStates.idle;
        }


        if (Input.GetKeyUp(KeyCode.W))
        {
            footSteps.Stop();
            curState = PlayerStates.idle;
        }
        if (Input.GetKeyUp(KeyCode.S))
        {
            footSteps.Stop();
            curState = PlayerStates.idle;
        }
        if (Input.GetKeyUp(KeyCode.D))
        {
            footSteps.Stop();
            curState = PlayerStates.idle;
        }
        if (Input.GetKeyUp(KeyCode.A))
        {
            footSteps.Stop();
            curState = PlayerStates.idle;
        }
        if (canClimb == true && Input.GetKeyDown(KeyCode.E))
        {
            curState = PlayerStates.climbing;
        }

    }
    void StateJetPack()
    {
        gravity = 0;
        onground = false;
        inair = false;
        canJump = false;
        CharacterController controller = GetComponent<CharacterController>();

        moveDirection = new Vector3(Input.GetAxis("Horizontal"), 0, Input.GetAxis("Vertical"));
        moveDirection = transform.TransformDirection(moveDirection);
        moveDirection *= moveSpeed;


        controller.Move(moveDirection * Time.deltaTime);

        if (Input.GetKey(KeyCode.W) && hasFuel == true)
        {

            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;

        }

        if (Input.GetKey(KeyCode.S) && hasFuel == true)
        {
            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;
        }

        if (Input.GetKey(KeyCode.A) && hasFuel == true)
        {
            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;

        }
        if (Input.GetKey(KeyCode.D) && hasFuel == true)
        {
            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;
        }
        if (Input.GetMouseButton(0) && hasFuel == true)
        {
            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;
            //gameObject.transform.localPosition += transform.up * sprintSpeed * Time.deltaTime;
            moveDirection.y = 40f * sprintSpeed * Time.deltaTime;
        }
        if (Input.GetMouseButton(1) && hasFuel == true)
        {
            UseFuel.CurFuel -= DecreasedAmount * Time.deltaTime;
            UseFuel.FuelBar.fillAmount -= DecreasedAmount * Time.deltaTime;
            //gameObject.transform.localPosition -= transform.up * sprintSpeed * Time.deltaTime;
            moveDirection.y = -40f * sprintSpeed * Time.deltaTime;
        }
        if (hasFuel == false)
        {
            zeroGravity = false;
            curState = PlayerStates.Inair;
        }
        if (inPast == false)
        {
            zeroGravity = false;
        }
        if (zeroGravity == false)
        {
            curState = PlayerStates.basicmovement;
        }

    }

    void StateClimbing()
    {
        if (zeroGravity == true)
        {
            curState = PlayerStates.jetPack;
        }
        //Debug.Log("You are in climbing State");
        if (Input.GetKey(KeyCode.W) && canClimb == true)
        {
            gameObject.transform.position += Vector3.up;
        }
        if (canClimb == false)
        {
            // climbText.SetActive (false);
            curState = PlayerStates.basicmovement;
        }
    }

    void StateInair()
    {
        inair = true;
        if (!flipped)
        {
            gravity = gravityValue;
        }
        if (flipped)
        {
            gravity = -gravityValue;
        }
        if (onground)
        {
            inair = false;
            curState = PlayerStates.idle;
        }
    }

    public bool GetFutureState()
    {
        return isFuture;
    }

    

    void OnTriggerEnter(Collider other)
    {
        if (other.tag == "Wall")
        {
            curState = PlayerStates.basicmovement;
        }
        if (other.name == "ClimbingTrigger")
        {
            canClimb = true;
            //climbText.SetActive (true);
            if (other.name == "GravityButtonController" && Input.GetKey(KeyCode.G) && hasFuel == false)
            {
                //Show You have No Fuel Panel 
            }
        }
    }

    void OnTriggerStay(Collider other)
    {
        if (other.name == "GravityButtonController" && Input.GetKey(KeyCode.G) && hasFuel == true)
        {
            zeroGravity = true;
        }
    }

    void OnTriggerExit(Collider other)
    {
        if (other.name == "ClimbingTrigger")
        {
            canClimb = false;
            //climbText.SetActive (false);
        }

    }

    public void groundtest()
    {

        curState = PlayerStates.idle;
    }

    public void airtest()
    {
        if (curState != PlayerStates.jetPack)
        {
            curState = PlayerStates.Inair;
        }
    }
}

