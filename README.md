# E2Metrix-Membrane-Fouling-Model

<h2> Description </h2> 

<p>The purpose of this simulation is to mathematically determine the fouling mechanism on MF/UF membranes during operation with one hour electrocoagulated thickener overflow water. The model is based off empirical data acquired during 4 hour experiments, with backflush at every hour. The model is currently implemented with a 200 nm TiO2 membrane, at 20, 40, and 60 psi. Here, the model will take experimental flux values during each hour of operation and optimize values for the constants used in each mechanism's equation by minimizing their residual square sum function. The K values for 7 different fouling mechanisms, including 3 combined models are calculated in the model: </p>

<p>1. Complete pore blocking </p>
<p>2. Intermediate pore blocking </p>
<p>3. Pore constriction </p>
<p>4. Cake filtration </p>
<p>5. Combined cake filtration and complete pore blocking </p>
<p>6. Combined intermediate blocking and cake filtration</p>
<p>7. Combined pore constriction and cake filtration </p>

<br>
<p>A second functionality of this model is to predict other parameters such as the cake resistance, cake thickness, and cake particle size based off other intrinsic membrane and fluid dynamics properties of the system. Once the constant Kc of the most optimal fouling mechanism is calculated by the previous section, it can then be used for determining these other attributes of our foulant. </p>

<br>
<h2> Examples of use </h2>
  
<p>The various functionalities of this model are described in the following video:.</p>

<br> 
<h2> Technologies </h2>

<p>Jupyter notebook was used for the model and ease of access.</p>
<p>Two VBA scripts are provided for processing the raw data from LABVIEW of bench MF/UF experiments, and processing the data from CIAL results for electrocoagulated samples.</p>
